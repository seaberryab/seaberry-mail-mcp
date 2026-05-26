using Microsoft.Identity.Client;
using Microsoft.Extensions.Logging;
using SeaberryMailMcp.Config;

namespace SeaberryMailMcp.Auth;

public sealed class TokenAcquirer
{
    public static readonly string[] Scopes =
    [
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/Calendars.ReadWrite",
        "https://graph.microsoft.com/Tasks.ReadWrite",
    ];

    private readonly AppConfig _config;
    private readonly ILogger<TokenAcquirer> _log;
    private IPublicClientApplication? _app;

    public TokenAcquirer(AppConfig config, ILogger<TokenAcquirer> log)
    {
        _config = config;
        _log = log;
    }

    private async Task<IPublicClientApplication> EnsureAppAsync()
    {
        if (_app is not null) return _app;

        if (!_config.IsProvisioned)
        {
            throw new InvalidOperationException(
                $"Config missing ClientId/TenantId at {AppConfig.ConfigFilePath}. {AppConfig.SetupScriptHint}");
        }

        _app = PublicClientApplicationBuilder
            .Create(_config.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, _config.TenantId)
            // CRITICAL: must be "http://localhost" (no port). MSAL picks a random port at runtime;
            // Entra's loopback rule accepts any port as a wildcard match.
            // Do NOT use .WithDefaultRedirectUri() — on .NET it resolves to the native-client URI
            // which fails interactive flow (no embedded WebView2).
            .WithRedirectUri("http://localhost")
            .Build();

        await CrossPlatCache.AttachAsync(_app).ConfigureAwait(false);
        return _app;
    }

    public sealed record AcquireOutcome(
        string AccessToken,
        string SignedInAs,
        bool DeviceCodeUsed,
        string? DeviceCodeMessage,
        IEnumerable<string> GrantedScopes);

    public async Task<AcquireOutcome> AcquireAsync(CancellationToken ct, bool force = false)
    {
        var app = await EnsureAppAsync().ConfigureAwait(false);

        if (force)
        {
            // Wipe every cached account so silent acquisition is guaranteed to fail and
            // we drop into the interactive consent flow. Use when scopes have changed.
            var existing = await app.GetAccountsAsync().ConfigureAwait(false);
            foreach (var acct in existing)
            {
                await app.RemoveAsync(acct).ConfigureAwait(false);
            }
            _log.LogInformation("force=true: removed {Count} cached account(s)", existing.Count());
        }

        // 1. Silent first.
        var accounts = await app.GetAccountsAsync().ConfigureAwait(false);
        var first = accounts.FirstOrDefault();
        if (first is not null)
        {
            try
            {
                var silent = await app.AcquireTokenSilent(Scopes, first).ExecuteAsync(ct).ConfigureAwait(false);
                return new AcquireOutcome(silent.AccessToken, silent.Account.Username, false, null, silent.Scopes);
            }
            catch (MsalUiRequiredException)
            {
                _log.LogInformation("silent token acquisition needs UI; falling through to interactive");
            }
        }

        // 2. Interactive (system browser, loopback).
        try
        {
            var interactive = await app.AcquireTokenInteractive(Scopes)
                .WithUseEmbeddedWebView(false)
                .ExecuteAsync(ct)
                .ConfigureAwait(false);
            return new AcquireOutcome(interactive.AccessToken, interactive.Account.Username, false, null, interactive.Scopes);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _log.LogWarning(ex, "interactive sign-in failed; falling back to device code");
        }

        // 3. Device code fallback (WSL, headless).
        string? deviceCodeMessage = null;
        var device = await app.AcquireTokenWithDeviceCode(Scopes, dc =>
        {
            deviceCodeMessage = dc.Message;
            // The MCP host captures stderr, but the user needs to see this URL+code.
            // Logging at Information so it appears in the host's log stream.
            _log.LogInformation("device code sign-in: {Message}", dc.Message);
            return Task.CompletedTask;
        }).ExecuteAsync(ct).ConfigureAwait(false);

        return new AcquireOutcome(device.AccessToken, device.Account.Username, true, deviceCodeMessage, device.Scopes);
    }
}
