using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using XlWright.Core;
using XlWright.Core.Com;
using XlWright.McpServer;

var options = ServerOptions.Parse(args, out var error);
if (options is null)
{
    Console.Error.WriteLine(error);
    Console.Error.WriteLine("Usage: XlWright.McpServer --project <repository> [--no-excel]");
    return 2;
}

// One server holds at most one session (docs/architecture.md "構成"). When stdin closes, the server ends
// without touching the Excel or ending it; the next start cleans up (stage 2).
using var session = options.NoExcel ? null : new ExcelSession(new ComExcelLauncher(), SessionTimeouts.Default);

// The empty builder reads no configuration files and adds no stdout logging.
// stdout carries only JSON-RPC frames, so logs go to stderr.
var builder = Host.CreateEmptyApplicationBuilder(settings: null);
builder.Logging.AddConsole(console => console.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Logging.SetMinimumLevel(LogLevel.Warning);
builder.Services
    .AddMcpServer(server => server.ServerInfo = new()
    {
        Name = ServerIdentity.ServerName,
        Version = ServerIdentity.Current.Version,
    })
    .WithStdioServerTransport()
    .WithTools<ServerInfoTool>()
    .WithTools(new SessionTools(options, session).Create())
    .WithTools(new RangeTools(session).Create())
    .WithTools(new VbaTools(session).Create());

await builder.Build().RunAsync();
return 0;
