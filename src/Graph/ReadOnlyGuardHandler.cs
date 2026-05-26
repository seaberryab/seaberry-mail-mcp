namespace SeaberryMailMcp.Graph;

/// Defense-in-depth: default-denies every outbound DELETE in the Graph HTTP pipeline,
/// with an explicit allowlist for paths where deletion is a deliberately-supported feature.
///
/// Current allowlist:
/// - DELETE /me/todo/lists/{listId}/tasks/{taskId}  — task-level deletion (delete_todo tool)
///
/// Notable deny-paths (NOT in the allowlist, and never should be without serious thought):
/// - DELETE /me/messages/{id}            — would delete an email
/// - DELETE /me/mailFolders/{id}         — would delete an entire mail folder
/// - DELETE /me/events/{id}              — would delete a calendar event (and notify attendees!)
/// - DELETE /me/todo/lists/{listId}      — would delete a whole task list (and every task in it)
///
/// To add a new permitted deletion path, edit AllowDelete below — that's the deliberate code
/// edit. The guard does not look at request bodies or query strings, only the URL path.
internal sealed class ReadOnlyGuardHandler : DelegatingHandler
{
    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        if (request.Method == HttpMethod.Delete && !IsAllowedDelete(request.RequestUri))
        {
            throw new InvalidOperationException(
                $"DELETE {request.RequestUri} blocked by ReadOnlyGuardHandler. " +
                "This MCP only permits DELETE on specifically allowlisted paths " +
                "(currently: /me/todo/lists/.../tasks/...). " +
                "To remove other items, do it in Outlook directly.");
        }
        return base.SendAsync(request, cancellationToken);
    }

    private static bool IsAllowedDelete(Uri? uri)
    {
        if (uri is null) return false;
        var path = uri.AbsolutePath;
        // Allow only task-level deletes under /me/todo/lists/.../tasks/...
        // Both segments required; this rejects DELETE /me/todo/lists/{id} (whole list).
        return path.Contains("/me/todo/lists/", StringComparison.OrdinalIgnoreCase)
            && path.Contains("/tasks/", StringComparison.OrdinalIgnoreCase);
    }
}
