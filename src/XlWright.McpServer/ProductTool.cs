using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using XlWright.Core;

namespace XlWright.McpServer;

// A subclass rather than McpServerTool.Create: the contract fixes the schemas, and arguments must reach
// ToolArguments unconverted so that invalid input fails in the contract's shape.
internal sealed class ProductTool(Tool tool, IReadOnlyCollection<string> arguments, ExcelSession session,
    Func<ToolArguments, Task<OperationResult>> run) : McpServerTool
{
    public override Tool ProtocolTool => tool;

    public override IReadOnlyList<object> Metadata => [];

    // Cancellation is ignored on purpose: an accepted COM call cannot be stopped, and the session's deadline decides.
    public override async ValueTask<CallToolResult> InvokeAsync(RequestContext<CallToolRequestParams> request, CancellationToken cancellationToken = default)
    {
        try { return ToolResponses.Success(await run(new(tool.Name, session, request.Params?.Arguments, arguments))); }
        catch (SessionOperationException failure) { return ToolResponses.Failure(failure); }
    }
}
