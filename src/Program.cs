using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SeaberryMailMcp.Auth;
using SeaberryMailMcp.Config;
using SeaberryMailMcp.Graph;

var builder = Host.CreateApplicationBuilder(args);

// CRITICAL for stdio MCP: stdout is the protocol channel. ALL logging must go to stderr,
// otherwise a stray log line corrupts the MCP transport and the client disconnects.
builder.Logging.ClearProviders();
builder.Logging.AddConsole(opts =>
{
    opts.LogToStandardErrorThreshold = LogLevel.Trace;
});

builder.Services.AddSingleton(_ => AppConfig.Load());
builder.Services.AddSingleton<TokenAcquirer>();
builder.Services.AddSingleton<GraphClientFactory>();

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
