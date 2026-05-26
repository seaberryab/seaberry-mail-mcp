using System.ComponentModel;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using SeaberryMailMcp.Graph;
using static SeaberryMailMcp.Graph.GraphOps;

namespace SeaberryMailMcp.Tools;

[McpServerToolType]
public sealed class MailTools
{
    private const int DefaultListTop = 25;
    private const int DefaultSearchTop = 25;
    private const int MaxPagedItems = 100;

    private static readonly string[] LeanSelect =
    [
        "id", "subject", "from", "toRecipients", "receivedDateTime",
        "isRead", "hasAttachments", "bodyPreview", "webLink",
    ];

    // ─── list_messages ───────────────────────────────────────────────────────

    [McpServerTool(Name = "list_messages")]
    [Description(
        "List recent messages in a mail folder (default: Inbox). Use for triage — " +
        "'what came in today/this week'. Returns lean metadata; call get_message for full body. " +
        "Ordered by receivedDateTime descending.")]
    public static async Task<object> ListMessages(
        GraphClientFactory factory,
        [Description("Mail folder name. Common: 'Inbox', 'Drafts', 'SentItems', 'Archive'. Default: Inbox.")]
        string? folder = null,
        [Description("Maximum number of messages to return. Default 25, max 100.")]
        int? top = null,
        [Description("ISO-8601 datetime; only return messages received on or after this. Optional.")]
        string? since = null,
        CancellationToken ct = default)
    {
        try
        {
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);
        var folderName = string.IsNullOrWhiteSpace(folder) ? "Inbox" : folder;
        var take = Math.Clamp(top ?? DefaultListTop, 1, MaxPagedItems);

        string? filter = null;
        if (!string.IsNullOrWhiteSpace(since))
        {
            // Graph $filter on receivedDateTime requires an Edm.DateTimeOffset literal:
            // ISO-8601 with a timezone marker (Z or ±HH:MM). A naked "2026-05-25T00:00:00"
            // (no zone) makes Graph return 400. Parse flexibly, AssumeLocal when the input
            // has no zone, convert to UTC, emit with Z.
            if (!System.DateTimeOffset.TryParse(
                    since,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.AssumeLocal,
                    out var parsed))
            {
                throw new ArgumentException(
                    $"could not parse 'since' as a datetime: '{since}'. " +
                    "Use ISO-8601, e.g. '2026-05-25' or '2026-05-25T00:00:00'.", nameof(since));
            }
            filter = $"receivedDateTime ge {parsed.ToUniversalTime():yyyy-MM-ddTHH:mm:ssZ}";
        }

        var response = await ExecAsync($"list messages in folder '{folderName}'",
            () => client.Me.MailFolders[folderName].Messages
                .GetAsync(rc =>
                {
                    rc.QueryParameters.Top = take;
                    rc.QueryParameters.Orderby = ["receivedDateTime desc"];
                    rc.QueryParameters.Select = LeanSelect;
                    if (filter is not null) rc.QueryParameters.Filter = filter;
                }, ct))
            .ConfigureAwait(false);

        var items = (response?.Value ?? []).Select(ToSummary).ToList();
        return new
        {
            folder = folderName,
            count = items.Count,
            messages = items,
        };
        }
        catch (Exception ex) { return FormatError("list_messages", ex); }
    }

    // ─── search_messages ─────────────────────────────────────────────────────

    [McpServerTool(Name = "search_messages")]
    [Description(
        "Search messages across the mailbox. Accepts plain text or KQL-style operators " +
        "(e.g. 'from:alice@example.com', 'subject:budget', 'hasAttachments:true', " +
        "'from:alice AND subject:Q3'). Note: $search cannot be combined with $orderby, " +
        "so results come back in relevance order. Paged automatically up to ~100 items.")]
    public static async Task<object> SearchMessages(
        GraphClientFactory factory,
        [Description("Search query. Plain text matches across subject/body/from. KQL operators supported.")]
        string query,
        [Description("Maximum items to return across paging. Default 25, max 100.")]
        int? top = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw new ArgumentException("query must be non-empty", nameof(query));
        }

        var client = await factory.CreateAsync(ct).ConfigureAwait(false);
        var take = Math.Clamp(top ?? DefaultSearchTop, 1, MaxPagedItems);

        // Graph requires $search values wrapped in double quotes.
        var quoted = $"\"{query.Replace("\"", "\\\"")}\"";

        // Action that adds ConsistencyLevel — required for $search. Must be re-applied
        // to every paged request because headers do NOT propagate by default.
        Action<Microsoft.Kiota.Abstractions.RequestInformation>? withConsistencyLevel = req =>
        {
            req.Headers.Add("ConsistencyLevel", "eventual");
        };

        var first = await ExecAsync("search messages (page 1)",
            () => client.Me.Messages
                .GetAsync(rc =>
                {
                    rc.QueryParameters.Search = quoted;
                    rc.QueryParameters.Top = take;
                    rc.QueryParameters.Select = LeanSelect;
                    rc.Headers.Add("ConsistencyLevel", "eventual");
                }, ct))
            .ConfigureAwait(false);

        var collected = new List<MessageSummary>();
        if (first?.Value is { } page1)
        {
            foreach (var m in page1)
            {
                if (collected.Count >= take) break;
                collected.Add(ToSummary(m));
            }
        }

        // Follow @odata.nextLink manually so we re-apply ConsistencyLevel on every page.
        var nextLink = first?.OdataNextLink;
        var pageIndex = 2;
        while (!string.IsNullOrWhiteSpace(nextLink) && collected.Count < take)
        {
            var capturedLink = nextLink;
            var capturedIndex = pageIndex;
            var nextResponse = await ExecAsync($"search messages (page {capturedIndex})",
                () => client.Me.Messages
                    .WithUrl(capturedLink)
                    .GetAsync(rc => rc.Headers.Add("ConsistencyLevel", "eventual"), ct))
                .ConfigureAwait(false);

            if (nextResponse?.Value is { } pageN)
            {
                foreach (var m in pageN)
                {
                    if (collected.Count >= take) break;
                    collected.Add(ToSummary(m));
                }
            }
            nextLink = nextResponse?.OdataNextLink;
            pageIndex++;
        }

        return new
        {
            query,
            count = collected.Count,
            messages = collected,
        };
        }
        catch (Exception ex) { return FormatError("search_messages", ex); }
    }

    // ─── get_message ─────────────────────────────────────────────────────────

    [McpServerTool(Name = "get_message")]
    [Description(
        "Fetch one message by id, including the full body. " +
        "Set bodyFormat='text' to have Graph extract plain text from HTML server-side.")]
    public static async Task<object> GetMessage(
        GraphClientFactory factory,
        [Description("Message id, as returned by list_messages or search_messages.")]
        string id,
        [Description("'text' (plain text body, server-extracted) or 'html' (raw HTML). Default 'text'.")]
        string? bodyFormat = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(id))
        {
            throw new ArgumentException("id must be non-empty", nameof(id));
        }

        var asText = !string.Equals(bodyFormat, "html", StringComparison.OrdinalIgnoreCase);
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var msg = await ExecAsync($"get message {id}",
            () => client.Me.Messages[id].GetAsync(rc =>
            {
                if (asText)
                {
                    // Graph extracts plain text from HTML for us — far more reliable than client-side stripping.
                    rc.Headers.Add("Prefer", "outlook.body-content-type=\"text\"");
                }
            }, ct))
            .ConfigureAwait(false);

        if (msg is null) throw new InvalidOperationException($"message {id} not found");

        return new
        {
            id = msg.Id,
            subject = msg.Subject,
            from = msg.From?.EmailAddress?.Address,
            fromName = msg.From?.EmailAddress?.Name,
            to = (msg.ToRecipients ?? []).Select(r => r.EmailAddress?.Address).ToList(),
            cc = (msg.CcRecipients ?? []).Select(r => r.EmailAddress?.Address).ToList(),
            receivedDateTime = msg.ReceivedDateTime,
            isRead = msg.IsRead,
            hasAttachments = msg.HasAttachments,
            bodyContentType = msg.Body?.ContentType?.ToString(),
            body = msg.Body?.Content,
            webLink = msg.WebLink,
        };
        }
        catch (Exception ex) { return FormatError("get_message", ex); }
    }

    // ─── create_draft ────────────────────────────────────────────────────────

    [McpServerTool(Name = "create_draft")]
    [Description(
        "Create a draft email in Outlook (lands in the Drafts folder). Does NOT send. " +
        "The user reviews and sends from Outlook manually. No domain allowlist is enforced on " +
        "draft recipients because drafts never leave the mailbox.")]
    public static async Task<object> CreateDraft(
        GraphClientFactory factory,
        [Description("Subject line.")]
        string subject,
        [Description("Body (HTML). Plain text also accepted; will be rendered as HTML if it contains tags.")]
        string body,
        [Description("Primary recipients (email addresses).")]
        string[] to,
        [Description("CC recipients (email addresses). Optional.")]
        string[]? cc = null,
        [Description("BCC recipients (email addresses). Optional.")]
        string[]? bcc = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(subject)) throw new ArgumentException("subject required", nameof(subject));
        if (to is null || to.Length == 0) throw new ArgumentException("to must have at least one recipient", nameof(to));

        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var msg = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = body ?? string.Empty,
            },
            ToRecipients = to.Select(ToRecipient).ToList(),
            CcRecipients = (cc ?? []).Select(ToRecipient).ToList(),
            BccRecipients = (bcc ?? []).Select(ToRecipient).ToList(),
        };

        var created = await ExecAsync("create draft",
            () => client.Me.Messages.PostAsync(msg, cancellationToken: ct))
            .ConfigureAwait(false);
        if (created is null) throw new InvalidOperationException("draft creation returned null");

        return new
        {
            status = "draft_created",
            id = created.Id,
            subject = created.Subject,
            webLink = created.WebLink,
            note = "Draft is in your Outlook Drafts folder. Open it via webLink to review and send.",
        };
        }
        catch (Exception ex) { return FormatError("create_draft", ex); }
    }

    // ─── helpers ─────────────────────────────────────────────────────────────

    private static Recipient ToRecipient(string addr) => new()
    {
        EmailAddress = new EmailAddress { Address = addr },
    };

    private sealed record MessageSummary(
        string? Id,
        string? Subject,
        string? FromAddress,
        string? FromName,
        DateTimeOffset? ReceivedDateTime,
        bool? IsRead,
        bool? HasAttachments,
        string? BodyPreview,
        string? WebLink);

    private static MessageSummary ToSummary(Message m) => new(
        m.Id,
        m.Subject,
        m.From?.EmailAddress?.Address,
        m.From?.EmailAddress?.Name,
        m.ReceivedDateTime,
        m.IsRead,
        m.HasAttachments,
        m.BodyPreview,
        m.WebLink);
}
