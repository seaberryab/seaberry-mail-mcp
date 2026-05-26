namespace SeaberryMailMcp.Graph;

/// Wrappers that turn Microsoft.Graph's ODataError (and any other tool exception)
/// into something useful at the MCP boundary.
///
/// Two distinct concerns:
///
/// 1. ExecAsync — turns ODataError's unhelpful default Exception.Message
///    ("An error occurred") into "Graph error during {op}: [{status}] {code}: {message}".
///    Every Graph call across all tool classes goes through this.
///
/// 2. FormatError — the MCP C# SDK's default exception handler returns the generic
///    "An error occurred invoking '{tool}'" to the client and drops the inner message.
///    To guarantee the diagnostic detail reaches Claude, every tool method catches
///    and RETURNS a structured error object instead of throwing.
internal static class GraphOps
{
    public static object FormatError(string toolName, Exception ex) => new
    {
        status = "error",
        tool = toolName,
        error = ex.Message,
        type = ex.GetType().Name,
        inner = ex.InnerException?.Message,
    };


    public static async Task<T> ExecAsync<T>(string opName, Func<Task<T>> op)
    {
        try
        {
            return await op().ConfigureAwait(false);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            throw Wrap(opName, ex);
        }
    }

    public static async Task ExecAsync(string opName, Func<Task> op)
    {
        try
        {
            await op().ConfigureAwait(false);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            throw Wrap(opName, ex);
        }
    }

    private static InvalidOperationException Wrap(string opName, Microsoft.Graph.Models.ODataErrors.ODataError ex)
    {
        var code = ex.Error?.Code ?? "(no code)";
        var msg = ex.Error?.Message ?? "(no message)";
        return new InvalidOperationException(
            $"Graph error during {opName}: [{ex.ResponseStatusCode}] {code}: {msg}", ex);
    }
}
