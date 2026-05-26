using System.ComponentModel;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using SeaberryMailMcp.Config;
using SeaberryMailMcp.Graph;
using static SeaberryMailMcp.Graph.GraphOps;

namespace SeaberryMailMcp.Tools;

[McpServerToolType]
public sealed class CalendarTools
{
    private const string DefaultTimeZone = "Europe/Stockholm";
    private const int MaxEvents = 250;

    // ─── list_events ─────────────────────────────────────────────────────────

    [McpServerTool(Name = "list_events")]
    [Description(
        "List calendar events in a date range. Uses /me/calendarView which expands recurring " +
        "series into individual instances within the window — preferred over /me/events for " +
        "'what's on my calendar this week'. Times returned in Europe/Stockholm by default.")]
    public static async Task<object> ListEvents(
        GraphClientFactory factory,
        [Description("Start of window, ISO-8601 (e.g. '2026-05-22T00:00:00'). Inclusive.")]
        string start,
        [Description("End of window, ISO-8601. Exclusive.")]
        string end,
        [Description("IANA time zone for inputs and returned times. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(start)) throw new ArgumentException("start required", nameof(start));
        if (string.IsNullOrWhiteSpace(end)) throw new ArgumentException("end required", nameof(end));

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var response = await ExecAsync($"list events {start}..{end}",
            () => client.Me.CalendarView
                .GetAsync(rc =>
                {
                    rc.QueryParameters.StartDateTime = start;
                    rc.QueryParameters.EndDateTime = end;
                    rc.QueryParameters.Top = MaxEvents;
                    rc.QueryParameters.Orderby = ["start/dateTime"];
                    rc.QueryParameters.Select =
                    [
                        "id", "subject", "start", "end", "location", "attendees",
                        "organizer", "isOnlineMeeting", "onlineMeeting", "webLink", "bodyPreview",
                    ];
                    // Ask Graph to return start/end in the requested TZ rather than UTC.
                    rc.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
                }, ct))
            .ConfigureAwait(false);

        var items = (response?.Value ?? []).Select(ToSummary).ToList();
        return new
        {
            start,
            end,
            timeZone = tz,
            count = items.Count,
            events = items,
        };
        }
        catch (Exception ex) { return FormatError("list_events", ex); }
    }

    // ─── search_events ───────────────────────────────────────────────────────

    [McpServerTool(Name = "search_events")]
    [Description(
        "Find calendar events by subject substring. Uses $filter contains() on the subject field. " +
        "Note: this is a substring match, not full-text; combine with a date range if results are noisy.")]
    public static async Task<object> SearchEvents(
        GraphClientFactory factory,
        [Description("Substring to find in the event subject (case-insensitive).")]
        string query,
        [Description("Optional ISO-8601 start of window to constrain the search.")]
        string? start = null,
        [Description("Optional ISO-8601 end of window to constrain the search.")]
        string? end = null,
        [Description("IANA time zone for returned times. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(query)) throw new ArgumentException("query required", nameof(query));

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        // Escape single quotes in the query (OData literal).
        var escaped = query.Replace("'", "''");

        // If a date range is given, prefer calendarView (recurring expansion); otherwise plain events.
        if (!string.IsNullOrWhiteSpace(start) && !string.IsNullOrWhiteSpace(end))
        {
            var response = await ExecAsync($"search events in window (subject contains '{query}')",
                () => client.Me.CalendarView
                    .GetAsync(rc =>
                    {
                        rc.QueryParameters.StartDateTime = start;
                        rc.QueryParameters.EndDateTime = end;
                        rc.QueryParameters.Top = MaxEvents;
                        rc.QueryParameters.Filter = $"contains(subject, '{escaped}')";
                        rc.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
                    }, ct))
                .ConfigureAwait(false);

            var items = (response?.Value ?? []).Select(ToSummary).ToList();
            return new { query, scope = "calendarView", start, end, timeZone = tz, count = items.Count, events = items };
        }
        else
        {
            var response = await ExecAsync($"search events (subject contains '{query}')",
                () => client.Me.Events
                    .GetAsync(rc =>
                    {
                        rc.QueryParameters.Top = MaxEvents;
                        rc.QueryParameters.Filter = $"contains(subject, '{escaped}')";
                        rc.Headers.Add("Prefer", $"outlook.timezone=\"{tz}\"");
                    }, ct))
                .ConfigureAwait(false);

            var items = (response?.Value ?? []).Select(ToSummary).ToList();
            return new { query, scope = "events", timeZone = tz, count = items.Count, events = items };
        }
        }
        catch (Exception ex) { return FormatError("search_events", ex); }
    }

    // ─── block_time ──────────────────────────────────────────────────────────

    [McpServerTool(Name = "block_time")]
    [Description(
        "Create a personal calendar event on your own calendar with NO attendees. " +
        "Use for blocking deep-work time, appointments, reminders. Structurally cannot send invites — " +
        "for events with attendees, use create_meeting instead.")]
    public static async Task<object> BlockTime(
        GraphClientFactory factory,
        [Description("Event title.")]
        string subject,
        [Description("Start datetime, ISO-8601 (e.g. '2026-05-23T14:00:00').")]
        string start,
        [Description("End datetime, ISO-8601.")]
        string end,
        [Description("IANA time zone for start/end. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        [Description("Optional body / description.")]
        string? body = null,
        [Description("Optional location (free text).")]
        string? location = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(subject)) throw new ArgumentException("subject required", nameof(subject));
        if (string.IsNullOrWhiteSpace(start)) throw new ArgumentException("start required", nameof(start));
        if (string.IsNullOrWhiteSpace(end)) throw new ArgumentException("end required", nameof(end));

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var evt = new Event
        {
            Subject = subject,
            Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
            End = new DateTimeTimeZone { DateTime = end, TimeZone = tz },
            Body = body is null ? null : new ItemBody { ContentType = BodyType.Text, Content = body },
            Location = location is null ? null : new Location { DisplayName = location },
            // attendees explicitly omitted — this tool cannot invite anyone.
        };

        var created = await ExecAsync("block time (create event without attendees)",
            () => client.Me.Events.PostAsync(evt, cancellationToken: ct))
            .ConfigureAwait(false);
        if (created is null) throw new InvalidOperationException("event creation returned null");

        return new
        {
            status = "blocked",
            id = created.Id,
            subject = created.Subject,
            start = created.Start,
            end = created.End,
            webLink = created.WebLink,
        };
        }
        catch (Exception ex) { return FormatError("block_time", ex); }
    }

    // ─── create_meeting ──────────────────────────────────────────────────────

    [McpServerTool(Name = "create_meeting")]
    [Description(
        "Create a calendar event WITH attendees. Sends meeting invitations to each attendee. " +
        "Two safety gates: (1) sendInvites must be explicitly true (defaults false → tool errors with a " +
        "preview), and (2) every attendee's email domain must be in AllowedInviteDomains in " +
        "~/.config/mail-mcp/config.json (defaults to [] on fresh installs — every attendee rejected until " +
        "you explicitly opt in). To allow a domain, edit config.json and restart the MCP.")]
    public static async Task<object> CreateMeeting(
        AppConfig config,
        GraphClientFactory factory,
        [Description("Event title.")]
        string subject,
        [Description("Start datetime, ISO-8601 (e.g. '2026-05-23T14:00:00').")]
        string start,
        [Description("End datetime, ISO-8601.")]
        string end,
        [Description("Attendee email addresses. Required and non-empty — use block_time if you have no attendees.")]
        string[] attendees,
        [Description("Must be explicitly true to actually send invites. Default false → tool returns a preview without writing to Graph.")]
        bool? sendInvites = null,
        [Description("IANA time zone for start/end. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        [Description("Optional meeting body / agenda.")]
        string? body = null,
        [Description("Optional location (free text).")]
        string? location = null,
        [Description("If true, Graph generates a Microsoft Teams join link. Default false.")]
        bool? isOnlineMeeting = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(subject)) throw new ArgumentException("subject required", nameof(subject));
        if (string.IsNullOrWhiteSpace(start)) throw new ArgumentException("start required", nameof(start));
        if (string.IsNullOrWhiteSpace(end)) throw new ArgumentException("end required", nameof(end));

        if (attendees is null || attendees.Length == 0)
        {
            return new
            {
                status = "error",
                error = "create_meeting requires at least one attendee. Use block_time for a personal calendar block with no invitees.",
            };
        }

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var preview = new
        {
            subject,
            start,
            end,
            timeZone = tz,
            attendees,
            location,
            body,
            isOnlineMeeting = isOnlineMeeting ?? false,
        };

        // GATE 1: explicit sendInvites confirmation.
        if (sendInvites is not true)
        {
            return new
            {
                status = "preview_only",
                error = "sendInvites is false (default). Review the preview below; if correct, re-invoke with sendInvites=true to actually send the meeting invitations.",
                preview,
            };
        }

        // GATE 2: domain allowlist.
        var allowed = config.AllowedInviteDomains
            .Select(d => d.Trim().ToLowerInvariant())
            .Where(d => d.Length > 0)
            .ToHashSet();

        var disallowed = attendees
            .Select(a => new { email = a, domain = ExtractDomain(a) })
            .Where(x => x.domain is null || !allowed.Contains(x.domain))
            .ToList();

        if (disallowed.Count > 0)
        {
            return new
            {
                status = "blocked_by_allowlist",
                error = "One or more attendees are outside the configured invite domain allowlist. No event was created.",
                disallowedAttendees = disallowed,
                allowedDomains = config.AllowedInviteDomains,
                howToWiden = $"Edit {AppConfig.ConfigFilePath} and add the desired domain(s) to AllowedInviteDomains, then restart the MCP.",
                preview,
            };
        }

        // Both gates passed — actually create the meeting (Graph will send invites).
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var evt = new Event
        {
            Subject = subject,
            Start = new DateTimeTimeZone { DateTime = start, TimeZone = tz },
            End = new DateTimeTimeZone { DateTime = end, TimeZone = tz },
            Body = body is null ? null : new ItemBody { ContentType = BodyType.Html, Content = body },
            Location = location is null ? null : new Location { DisplayName = location },
            Attendees = attendees.Select(a => new Attendee
            {
                EmailAddress = new EmailAddress { Address = a },
                Type = AttendeeType.Required,
            }).ToList(),
            IsOnlineMeeting = isOnlineMeeting ?? false,
            OnlineMeetingProvider = (isOnlineMeeting ?? false) ? OnlineMeetingProviderType.TeamsForBusiness : null,
        };

        var created = await ExecAsync("create meeting (sends invites to attendees)",
            () => client.Me.Events.PostAsync(evt, cancellationToken: ct))
            .ConfigureAwait(false);
        if (created is null) throw new InvalidOperationException("event creation returned null");

        return new
        {
            status = "meeting_sent",
            id = created.Id,
            subject = created.Subject,
            start = created.Start,
            end = created.End,
            webLink = created.WebLink,
            attendees = (created.Attendees ?? []).Select(a => a.EmailAddress?.Address).ToList(),
            onlineMeetingJoinUrl = created.OnlineMeeting?.JoinUrl,
        };
        }
        catch (Exception ex) { return FormatError("create_meeting", ex); }
    }

    // NOTE: if update_event is added later, the same two gates (sendInvites + allowlist) apply.
    // PATCH /me/events/{id} with new attendees also fires invites.

    // ─── helpers ─────────────────────────────────────────────────────────────

    private static string? ExtractDomain(string email)
    {
        if (string.IsNullOrWhiteSpace(email)) return null;
        var at = email.LastIndexOf('@');
        if (at < 0 || at == email.Length - 1) return null;
        return email[(at + 1)..].Trim().ToLowerInvariant();
    }

    private sealed record EventSummary(
        string? Id,
        string? Subject,
        DateTimeTimeZone? Start,
        DateTimeTimeZone? End,
        string? LocationDisplayName,
        string? OrganizerAddress,
        List<string?>? AttendeeAddresses,
        bool? IsOnlineMeeting,
        string? OnlineMeetingJoinUrl,
        string? BodyPreview,
        string? WebLink);

    private static EventSummary ToSummary(Event e) => new(
        e.Id,
        e.Subject,
        e.Start,
        e.End,
        e.Location?.DisplayName,
        e.Organizer?.EmailAddress?.Address,
        (e.Attendees ?? []).Select(a => a.EmailAddress?.Address).ToList(),
        e.IsOnlineMeeting,
        e.OnlineMeeting?.JoinUrl,
        e.BodyPreview,
        e.WebLink);
}
