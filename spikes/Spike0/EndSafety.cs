namespace Spike0;

internal static class EndSafety
{
    internal static void RequireIdentity(bool matches)
    {
        if (!matches) throw new InvalidOperationException("EXCEL_NEEDS_REVIEW: process identity changed.");
    }

    internal static void RequireBooks(string managedPath, IEnumerable<string> paths, int protectedCount,
        int workbookCount, bool requireEmpty)
    {
        if (protectedCount != 0 || requireEmpty && workbookCount != 0 || paths.Any(path =>
            !StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(path), Path.GetFullPath(managedPath))))
            throw new EndRefusedException();
    }

    internal static void CloseAndQuit(Action<bool> verify, Action close, Action quit)
    {
        verify(false);
        close();
        // A second observation is required even when the first one allowed Close.
        verify(true);
        quit();
    }

    internal static void Fault(ExecutionGate gate, bool generationStarted, bool closeStarted, Exception error) =>
        gate.Fault(generationStarted, !closeStarted && (error is EndRefusedException or SaveFailedException));
}
