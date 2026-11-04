using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using XlWright.Core;
using XlWright.Core.Com;
using XlWright.McpServer;

// The show command for the person runs instead of the server (docs/api.md "表示要求のコマンド").
if (args is ["show", .. var showArgs])
{
    var app = ShowOptions.Parse(showArgs, out var showError);
    if (app is null)
    {
        Console.Error.WriteLine(showError);
        Console.Error.WriteLine(ShowOptions.Usage);
        return 2;
    }

    return new ExcelShow(new WindowsProcessQuery(), new WorkbookWindows(), Console.In, Console.Out, SessionTimeouts.Default.TerminateExit).Run(app);
}

var options = ServerOptions.Parse(args, out var error);
if (options is null)
{
    Console.Error.WriteLine(error);
    Console.Error.WriteLine("Usage: XlWright.McpServer --project <repository> [--no-excel]");
    return 2;
}

// One server holds at most one session (docs/architecture.md "構成"). When stdin closes, the server ends
// without touching the Excel or ending it; the next start cleans up (stage 2). The session is not disposed,
// so the app lock is closed by the OS when the process ends, never by a running server ("実行時の状態と所有記録").
var session = options.NoExcel ? null : new ExcelSession(new ComExcelLauncher(), new WindowsProcessQuery(), new WindowConnector(), SessionTimeouts.Default);

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
