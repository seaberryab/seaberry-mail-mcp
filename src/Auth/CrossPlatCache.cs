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
        helper.RegisterCache(app.UserTokenCache);
    }
}
