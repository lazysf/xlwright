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

        try
        {
            var options = new McpServerOptions
            {
                ServerInfo = new() { Name = "xlwright", Version = "0.1.0" }
            };
            if (args.Length != 0)
            {
                options.ToolCollection = new();
                options.ToolCollection.Add(McpServerTool.Create(new DevelopmentTools().Status, new()
                {
                    Name = "dev_status", ReadOnly = true, Idempotent = true, Destructive = false, OpenWorld = false
                }));
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
    }

    internal static string BuildHash => Convert.ToHexString(SHA256.HashData(
        File.ReadAllBytes(Assembly.GetExecutingAssembly().Location)));
}
