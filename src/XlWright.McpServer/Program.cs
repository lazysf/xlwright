using System.Reflection;
using System.Security.Cryptography;
using ModelContextProtocol.Server;

namespace XlWright.McpServer;

internal static class Program
{
    internal static async Task<int> Main(string[] args)
    {
        if (args.Length != 0 && args is not ["--development"])
        {
            Console.Error.WriteLine("Usage: XlWright.McpServer [--development]");
            return 1;
        }

        DevelopmentTools? development = null;
        try
        {
            if (args.Length != 0) development = new DevelopmentTools();
            var options = new McpServerOptions
            {
                ServerInfo = new() { Name = "xlwright", Version = "0.1.0" }
            };
            if (development is not null)
            {
                options.ToolCollection = new();
                options.ToolCollection.Add(McpServerTool.Create(development.Status,
                    new() { Name = "dev_status", ReadOnly = true }));
                options.ToolCollection.Add(McpServerTool.Create(development.Start,
                    new() { Name = "dev_excel_smoke_start", ReadOnly = false }));
                options.ToolCollection.Add(McpServerTool.Create(development.StartEndRefusal,
                    new() { Name = "dev_excel_end_refusal_start", ReadOnly = false }));
                options.ToolCollection.Add(McpServerTool.Create(development.SmokeStatus,
                    new() { Name = "dev_excel_smoke_status", ReadOnly = true }));
            }
            await using var server = ModelContextProtocol.Server.McpServer.Create(
                new StdioServerTransport(options), options);
            await server.RunAsync();
            return 0;
        }
        catch (Exception error)
        {
            // Messages can contain personal paths; detailed development output stays local.
            Console.Error.WriteLine($"xlwright server failed: {error.GetType().Name}");
            return 1;
        }
        finally
        {
            if (development is not null) await development.DisposeAsync();
        }
    }

    internal static string BuildHash => Convert.ToHexString(SHA256.HashData(
        File.ReadAllBytes(Assembly.GetExecutingAssembly().Location)));
}
