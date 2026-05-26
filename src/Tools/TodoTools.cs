using System.ComponentModel;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using SeaberryMailMcp.Graph;
using static SeaberryMailMcp.Graph.GraphOps;

namespace SeaberryMailMcp.Tools;

[McpServerToolType]
public sealed class TodoTools
{
    private const string DefaultTimeZone = "Europe/Stockholm";
    private const int DefaultListTop = 50;
    private const int MaxListTop = 200;

    // ─── list_todo_lists ─────────────────────────────────────────────────────

    [McpServerTool(Name = "list_todo_lists")]
    [Description(
        "List the user's To Do lists. Each list has its own id; you need a listId to read or " +
        "write tasks. The built-in default list has wellknownListName='defaultList' and is " +
        "what 'Tasks' refers to in Outlook.")]
    public static async Task<object> ListTodoLists(
        GraphClientFactory factory,
        CancellationToken ct = default)
    {
        try
        {
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);
        var lists = await GraphOps.ExecAsync("list todo lists",
            () => client.Me.Todo.Lists.GetAsync(cancellationToken: ct))
            .ConfigureAwait(false);

        var items = (lists?.Value ?? []).Select(l => new
        {
            id = l.Id,
            displayName = l.DisplayName,
            wellknownListName = l.WellknownListName?.ToString(),
            isOwner = l.IsOwner,
            isShared = l.IsShared,
        }).ToList();

        return new { count = items.Count, lists = items };
        }
        catch (Exception ex) { return FormatError("list_todo_lists", ex); }
    }

    // ─── list_todos ──────────────────────────────────────────────────────────

    [McpServerTool(Name = "list_todos")]
    [Description(
        "List tasks in a To Do list. By default queries the built-in 'Tasks' list and shows only " +
        "open tasks (not yet completed). Pass status='all' to include completed; status='completed' " +
        "for only completed. Ordered by due date when set, else by creation date.")]
    public static async Task<object> ListTodos(
        GraphClientFactory factory,
        [Description("Task list id (from list_todo_lists). If omitted, uses the built-in default list ('Tasks').")]
        string? listId = null,
        [Description("Status filter: 'open' (default, excludes completed), 'all', 'completed', 'notStarted', 'inProgress', 'waitingOnOthers', 'deferred'.")]
        string? status = null,
        [Description("Max number of tasks to return. Default 50, max 200.")]
        int? top = null,
        CancellationToken ct = default)
    {
        try
        {
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);
        var resolvedListId = await ResolveListIdAsync(client, listId, ct).ConfigureAwait(false);
        var take = Math.Clamp(top ?? DefaultListTop, 1, MaxListTop);

        var filter = BuildStatusFilter(status);

        var response = await GraphOps.ExecAsync($"list tasks in {resolvedListId}",
            () => client.Me.Todo.Lists[resolvedListId].Tasks
                .GetAsync(rc =>
                {
                    rc.QueryParameters.Top = take;
                    if (filter is not null) rc.QueryParameters.Filter = filter;
                    // Note: $orderby on dueDateTime can fail when many tasks have no due date.
                    // We sort client-side instead — open tasks with dueDateTime ascending, then the rest.
                }, ct))
            .ConfigureAwait(false);

        var items = (response?.Value ?? [])
            .Select(ToSummary)
            .OrderBy(t => t.DueDateTime?.DateTime ?? "9999")
            .ThenByDescending(t => t.CreatedDateTime)
            .ToList();

        return new
        {
            listId = resolvedListId,
            filter = filter ?? "(none)",
            count = items.Count,
            tasks = items,
        };
        }
        catch (Exception ex) { return FormatError("list_todos", ex); }
    }

    // ─── create_todo ─────────────────────────────────────────────────────────

    [McpServerTool(Name = "create_todo")]
    [Description(
        "Create a new task in a To Do list. If no listId is given, the task lands in the built-in " +
        "'Tasks' list. dueDate is optional — pass a date ('2026-05-30') or full ISO datetime.")]
    public static async Task<object> CreateTodo(
        GraphClientFactory factory,
        [Description("Task title (required).")]
        string title,
        [Description("Optional body / notes (plain text).")]
        string? body = null,
        [Description("Optional due date. Accepts date-only ('2026-05-30') or full ISO ('2026-05-30T17:00:00').")]
        string? dueDate = null,
        [Description("Importance: 'low', 'normal', 'high'. Default 'normal'.")]
        string? importance = null,
        [Description("Task list id (from list_todo_lists). If omitted, uses the default 'Tasks' list.")]
        string? listId = null,
        [Description("IANA time zone for the due date. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("title required", nameof(title));

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);
        var resolvedListId = await ResolveListIdAsync(client, listId, ct).ConfigureAwait(false);

        var task = new TodoTask
        {
            Title = title,
            Body = body is null ? null : new ItemBody { ContentType = BodyType.Text, Content = body },
            DueDateTime = ParseDue(dueDate, tz),
            Importance = ParseImportance(importance),
        };

        var created = await GraphOps.ExecAsync("create todo",
            () => client.Me.Todo.Lists[resolvedListId].Tasks
                .PostAsync(task, cancellationToken: ct))
            .ConfigureAwait(false);

        if (created is null) throw new InvalidOperationException("todo creation returned null");

        return new
        {
            status = "created",
            listId = resolvedListId,
            task = ToSummary(created),
        };
        }
        catch (Exception ex) { return FormatError("create_todo", ex); }
    }

    // ─── update_todo ─────────────────────────────────────────────────────────

    [McpServerTool(Name = "update_todo")]
    [Description(
        "Patch fields on an existing task. Only fields you pass are updated — others are left " +
        "untouched. Both listId AND id are required (a task id is only unique within its list).")]
    public static async Task<object> UpdateTodo(
        GraphClientFactory factory,
        [Description("Task list id.")]
        string listId,
        [Description("Task id (from list_todos).")]
        string id,
        [Description("New title. Optional.")]
        string? title = null,
        [Description("New body / notes (plain text). Optional. Pass an empty string to clear.")]
        string? body = null,
        [Description("New due date. Accepts date-only or full ISO. Optional. Pass 'clear' to remove the due date.")]
        string? dueDate = null,
        [Description("New status: 'notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'. Optional.")]
        string? status = null,
        [Description("New importance: 'low', 'normal', 'high'. Optional.")]
        string? importance = null,
        [Description("IANA time zone for the due date. Default 'Europe/Stockholm'.")]
        string? timeZone = null,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(listId)) throw new ArgumentException("listId required", nameof(listId));
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("id required", nameof(id));

        var tz = string.IsNullOrWhiteSpace(timeZone) ? DefaultTimeZone : timeZone;
        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        // Build a patch payload with ONLY the fields that were explicitly passed.
        var patch = new TodoTask();
        if (title is not null) patch.Title = title;
        if (body is not null) patch.Body = new ItemBody { ContentType = BodyType.Text, Content = body };
        if (dueDate is not null)
        {
            patch.DueDateTime = string.Equals(dueDate, "clear", StringComparison.OrdinalIgnoreCase)
                ? null
                : ParseDue(dueDate, tz);
        }
        if (status is not null) patch.Status = ParseStatus(status);
        if (importance is not null) patch.Importance = ParseImportance(importance);

        var updated = await GraphOps.ExecAsync("update todo",
            () => client.Me.Todo.Lists[listId].Tasks[id]
                .PatchAsync(patch, cancellationToken: ct))
            .ConfigureAwait(false);

        if (updated is null) throw new InvalidOperationException("todo update returned null");

        return new
        {
            status = "updated",
            listId,
            task = ToSummary(updated),
        };
        }
        catch (Exception ex) { return FormatError("update_todo", ex); }
    }

    // ─── complete_todo ───────────────────────────────────────────────────────

    [McpServerTool(Name = "complete_todo")]
    [Description(
        "Mark a task as completed. Convenience for the common 'mark X done' case — equivalent to " +
        "update_todo with status='completed'. Both listId AND id are required.")]
    public static async Task<object> CompleteTodo(
        GraphClientFactory factory,
        [Description("Task list id.")]
        string listId,
        [Description("Task id (from list_todos).")]
        string id,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(listId)) throw new ArgumentException("listId required", nameof(listId));
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("id required", nameof(id));

        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        var updated = await GraphOps.ExecAsync("complete todo",
            () => client.Me.Todo.Lists[listId].Tasks[id]
                .PatchAsync(new TodoTask { Status = Microsoft.Graph.Models.TaskStatus.Completed }, cancellationToken: ct))
            .ConfigureAwait(false);

        if (updated is null) throw new InvalidOperationException("todo complete returned null");

        return new
        {
            status = "completed",
            listId,
            task = ToSummary(updated),
        };
        }
        catch (Exception ex) { return FormatError("complete_todo", ex); }
    }

    // ─── delete_todo ─────────────────────────────────────────────────────────

    [McpServerTool(Name = "delete_todo")]
    [Description(
        "Permanently delete a task by id. NOT recoverable — the task is gone, not moved to a trash " +
        "list. Pass an explicit id (no bulk-delete-by-query). Both listId AND id are required. " +
        "ReadOnlyGuardHandler permits DELETE only on the /me/todo/lists/.../tasks/... path.")]
    public static async Task<object> DeleteTodo(
        GraphClientFactory factory,
        [Description("Task list id.")]
        string listId,
        [Description("Task id (from list_todos).")]
        string id,
        CancellationToken ct = default)
    {
        try
        {
        if (string.IsNullOrWhiteSpace(listId)) throw new ArgumentException("listId required", nameof(listId));
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("id required", nameof(id));

        var client = await factory.CreateAsync(ct).ConfigureAwait(false);

        await GraphOps.ExecAsync("delete todo",
            () => client.Me.Todo.Lists[listId].Tasks[id]
                .DeleteAsync(cancellationToken: ct))
            .ConfigureAwait(false);

        return new
        {
            status = "deleted",
            listId,
            id,
        };
        }
        catch (Exception ex) { return FormatError("delete_todo", ex); }
    }

    // ─── helpers ─────────────────────────────────────────────────────────────

    private static async Task<string> ResolveListIdAsync(
        Microsoft.Graph.GraphServiceClient client,
        string? listId,
        CancellationToken ct)
    {
        if (!string.IsNullOrWhiteSpace(listId)) return listId;

        // Fetch all lists and filter client-side. $filter on the enum-typed
        // wellknownListName property has been flaky in some Graph versions;
        // client-side filtering is simpler and the list count is small (usually <10).
        var lists = await GraphOps.ExecAsync("list todo lists (for default resolution)",
            () => client.Me.Todo.Lists.GetAsync(cancellationToken: ct))
            .ConfigureAwait(false);

        var all = lists?.Value ?? [];
        var def = all.FirstOrDefault(l => l.WellknownListName == WellknownListName.DefaultList);

        if (def?.Id is not null) return def.Id;

        // Fallback: if the default list isn't flagged for any reason, prefer any list
        // named exactly "Tasks" (the conventional default display name).
        var byName = all.FirstOrDefault(l =>
            string.Equals(l.DisplayName, "Tasks", StringComparison.OrdinalIgnoreCase));
        if (byName?.Id is not null) return byName.Id;

        if (all.Count == 0)
        {
            throw new InvalidOperationException(
                "no To Do lists found on this account. Open Microsoft To Do " +
                "(https://to-do.office.com or the Tasks pane in Outlook on the web) once to initialize " +
                "the default list, then retry.");
        }

        throw new InvalidOperationException(
            $"could not identify a default To Do list among {all.Count} list(s). " +
            "Call list_todo_lists to inspect them and pass listId explicitly.");
    }

    // Error-surfacing wrapper lifted to SeaberryMailMcp.Graph.GraphOps so the same
    // ODataError unwrapping is applied uniformly across mail, calendar, and todo tools.

    private static string? BuildStatusFilter(string? status)
    {
        if (string.IsNullOrWhiteSpace(status) || string.Equals(status, "open", StringComparison.OrdinalIgnoreCase))
        {
            // Default: hide completed.
            return "status ne 'completed'";
        }
        if (string.Equals(status, "all", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }
        // Specific status value — Graph stores them as camelCase strings.
        var normalized = char.ToLowerInvariant(status[0]) + status[1..];
        return $"status eq '{normalized}'";
    }

    private static DateTimeTimeZone? ParseDue(string? input, string tz)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        // Accept date-only ("2026-05-30") or full ISO ("2026-05-30T17:00:00").
        // Graph wants a DateTime string + a separate TimeZone field.
        var dt = input.Contains('T') ? input : $"{input}T00:00:00";
        return new DateTimeTimeZone { DateTime = dt, TimeZone = tz };
    }

    private static Microsoft.Graph.Models.TaskStatus? ParseStatus(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        return input.ToLowerInvariant() switch
        {
            "notstarted" or "not_started" or "open" => Microsoft.Graph.Models.TaskStatus.NotStarted,
            "inprogress" or "in_progress" or "doing" => Microsoft.Graph.Models.TaskStatus.InProgress,
            "completed" or "done" => Microsoft.Graph.Models.TaskStatus.Completed,
            "waitingonothers" or "waiting" or "blocked" => Microsoft.Graph.Models.TaskStatus.WaitingOnOthers,
            "deferred" or "snoozed" => Microsoft.Graph.Models.TaskStatus.Deferred,
            _ => throw new ArgumentException(
                $"unknown status '{input}'. valid: notStarted, inProgress, completed, waitingOnOthers, deferred"),
        };
    }

    private static Microsoft.Graph.Models.Importance? ParseImportance(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        return input.ToLowerInvariant() switch
        {
            "low" => Microsoft.Graph.Models.Importance.Low,
            "normal" or "medium" => Microsoft.Graph.Models.Importance.Normal,
            "high" => Microsoft.Graph.Models.Importance.High,
            _ => throw new ArgumentException($"unknown importance '{input}'. valid: low, normal, high"),
        };
    }

    private sealed record TodoSummary(
        string? Id,
        string? Title,
        string? Status,
        string? Importance,
        DateTimeTimeZone? DueDateTime,
        DateTimeTimeZone? CompletedDateTime,
        DateTimeOffset? CreatedDateTime,
        DateTimeOffset? LastModifiedDateTime,
        string? BodyPreview);

    private static TodoSummary ToSummary(TodoTask t)
    {
        // bodyPreview-equivalent: trim the body content to 200 chars.
        var preview = t.Body?.Content;
        if (preview is not null && preview.Length > 200) preview = preview[..200] + "…";

        return new TodoSummary(
            t.Id,
            t.Title,
            t.Status?.ToString(),
            t.Importance?.ToString(),
            t.DueDateTime,
            t.CompletedDateTime,
            t.CreatedDateTime,
            t.LastModifiedDateTime,
            preview);
    }
}
