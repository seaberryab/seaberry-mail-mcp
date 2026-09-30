using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using SeaberryMailMcp.Config;

namespace SeaberryMailMcp.Auth;

internal static class CrossPlatCache
{
    private const string CacheFileName = "msal.cache.bin";
    private const string KeyChainServiceName = "seaberry-mail-mcp";
    private const string KeyChainAccountName = "msal-token-cache";
    private const string LinuxKeyRingSchema = "io.github.seaberryab.mailmcp.tokencache";
    private const string LinuxKeyRingCollection = "default";
    private const string LinuxKeyRingLabel = "Seaberry Mail MCP MSAL token cache";

    // Opt-in: on Linux without a usable keyring, fall back to a user-only (0600) cache file.
    private const string AllowFileCacheEnvVar = "MAIL_MCP_ALLOW_FILE_CACHE";

    public static async Task AttachAsync(IPublicClientApplication app)
    {
        Directory.CreateDirectory(AppConfig.ConfigDirectory);

        var storageProps = new StorageCreationPropertiesBuilder(CacheFileName, AppConfig.ConfigDirectory)
            .WithLinuxKeyring(
                LinuxKeyRingSchema,
                LinuxKeyRingCollection,
                LinuxKeyRingLabel,
                new KeyValuePair<string, string>("Version", "1"),
                new KeyValuePair<string, string>("Product", "seaberry-mail-mcp"))
            .WithMacKeyChain(KeyChainServiceName, KeyChainAccountName)
            .WithCacheChangedEvent(app.AppConfig.ClientId)
            .Build();

        var helper = await MsalCacheHelper.CreateAsync(storageProps).ConfigureAwait(false);

        if (OperatingSystem.IsLinux() && Environment.GetEnvironmentVariable(AllowFileCacheEnvVar) == "1")
        {
            try
            {
                helper.VerifyPersistence();
            }
            catch (MsalCachePersistenceException)
            {
                var fileProps = new StorageCreationPropertiesBuilder(CacheFileName, AppConfig.ConfigDirectory)
                    .WithLinuxUnprotectedFile()
                    .WithCacheChangedEvent(app.AppConfig.ClientId)
                    .Build();
                helper = await MsalCacheHelper.CreateAsync(fileProps).ConfigureAwait(false);
                RestrictToOwner(fileProps.CacheFilePath);
            }
        }

        helper.RegisterCache(app.UserTokenCache);
    }

    [System.Runtime.Versioning.UnsupportedOSPlatform("windows")]
    private static void RestrictToOwner(string path)
    {
        if (!File.Exists(path))
        {
            using (File.Create(path)) { }
        }
        File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);
    }
}
