using System.ComponentModel;
using System.Text.Json;

namespace XlWright.McpServer;

// Development mode only identifies the running build. Experiments run from the console.
internal sealed class DevelopmentTools
{
    private readonly string invocation = Guid.NewGuid().ToString("N");

    [Description("Read the identity of the running development server build. Does not access Excel.")]
    internal string Status() => JsonSerializer.Serialize(new { invocation, build = Program.BuildHash });
}
