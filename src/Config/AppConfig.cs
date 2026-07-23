using System.Text.Json;
using System.Text.Json.Serialization;

namespace SeaberryMailMcp.Config;

public sealed record AppConfig
{
    public string ClientId { get; init; } = string.Empty;
    public string TenantId { get; init; } = string.Empty;
    // Default empty — create_meeting will reject every attendee until the user opts in
    // explicitly by editing AllowedInviteDomains in ~/.config/mail-mcp/config.json.
    // Fail-loud-by-default is the whole point of the gate.
    public List<string> AllowedInviteDomains { get; init; } = [];

    // Display name of the folder used as a soft-archive destination by archive_messages.
    // Self-documenting in Outlook so a human reviewing the mailbox can see which messages
    // were moved by AI vs. by the user. Configurable so the convention can evolve.
    public string ArchiveFolderName { get; init; } = "Archived by AI";

    [JsonIgnore]
    public bool IsProvisioned => !string.IsNullOrWhiteSpace(ClientId) && !string.IsNullOrWhiteSpace(TenantId);

    public static string ConfigDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config", "mail-mcp");

    public static string ConfigFilePath => Path.Combine(ConfigDirectory, "config.json");

    public static string SetupScriptHint =>
        "Run ./setup.sh from your shell (signed in via `az login` as tenant admin) before invoking this MCP. " +
        "The MCP does not provision app registrations itself.";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
    };

    public static AppConfig Load()
    {
        if (!File.Exists(ConfigFilePath))
        {
            return new AppConfig();
        }

        var json = File.ReadAllText(ConfigFilePath);
        return JsonSerializer.Deserialize<AppConfig>(json, JsonOptions) ?? new AppConfig();
    }
}
