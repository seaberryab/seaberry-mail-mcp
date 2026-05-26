using System.ComponentModel;
using ModelContextProtocol.Server;
using SeaberryMailMcp.Auth;
using SeaberryMailMcp.Config;
using SeaberryMailMcp.Graph;

namespace SeaberryMailMcp.Tools;

[McpServerToolType]
public sealed class SetupTools
{
    [McpServerTool(Name = "ensure_setup")]
    [Description(
        "Idempotent auth bootstrap. Verifies that the Entra app registration has been provisioned " +
        "(via setup.sh) and that an MSAL token cache covers the requested scopes (Mail.ReadWrite, " +
        "Calendars.ReadWrite, Tasks.ReadWrite). If no cached account is found, triggers interactive " +
        "sign-in (or device-code fallback for WSL/headless). Does NOT provision Azure resources — " +
        "run ./setup.sh from your shell first if config is missing. Safe to re-run; returns immediately " +
        "when already signed in. Pass force=true to discard any cached token and force a fresh sign-in " +
        "(use this after scopes have been added/changed).")]
    public static async Task<object> EnsureSetup(
        AppConfig config,
        TokenAcquirer tokens,
        [Description("If true, discard any cached MSAL account and force a fresh interactive sign-in. Default false.")]
        bool? force = null,
        CancellationToken ct = default)
    {
        try
        {
            if (!config.IsProvisioned)
            {
                return new
                {
                    status = "not_provisioned",
                    error = AppConfig.SetupScriptHint,
                    configFilePath = AppConfig.ConfigFilePath,
                };
            }

            var outcome = await tokens.AcquireAsync(ct, force ?? false).ConfigureAwait(false);

            // MSAL can return scopes either as short names ("Mail.ReadWrite") or
            // fully-qualified ("https://graph.microsoft.com/Mail.ReadWrite") depending
            // on the IdP response. Normalize both sides to bare scope names before comparing.
            static string Bare(string s) => s.Replace("https://graph.microsoft.com/", "", StringComparison.OrdinalIgnoreCase);

            var requested = TokenAcquirer.Scopes.Select(Bare).ToList();
            var granted = outcome.GrantedScopes.Select(Bare).ToList();
            var missing = requested
                .Where(r => !granted.Any(g => string.Equals(g, r, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            return new
            {
                status = missing.Count == 0 ? "ready" : "scopes_missing",
                clientId = config.ClientId,
                tenantId = config.TenantId,
                signedInAs = outcome.SignedInAs,
                deviceCodeUsed = outcome.DeviceCodeUsed,
                deviceCodeMessage = outcome.DeviceCodeMessage,
                allowedInviteDomains = config.AllowedInviteDomains,
                requestedScopes = requested,
                grantedScopes = granted,
                missingScopes = missing,
                hint = missing.Count == 0
                    ? null
                    : "Some scopes are missing from the token. Confirm admin consent has been granted for them in Entra, then re-invoke with force=true.",
            };
        }
        catch (Exception ex)
        {
            return GraphOps.FormatError("ensure_setup", ex);
        }
    }
}
