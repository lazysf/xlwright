namespace XlWright.Core;

internal static class EndSafety
{
    // A changed identity is EXCEL_NEEDS_REVIEW like a refusal, but it isolates the session, so it is not EndRefusedException.
    internal static void RequireIdentity(bool matches)
    {
        if (!matches)
            throw new CoreException(ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview, "Excel process identity changed. Do not resend; follow manual recovery.");
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

    // A clear Excel-side failure or absence keeps a ready session ready; anything else may have left unknown state.
    internal static void Fault(ExecutionGate gate, bool generationStarted, bool closeStarted, Exception error) =>
        gate.Fault(generationStarted, !closeStarted &&
            error is EndRefusedException or CoreException { Code: ErrorCode.ExcelError or ErrorCode.NotFound });
}

// Raised only when an unmanaged workbook or Protected View is confirmed, which keeps a ready session ready.
internal sealed class EndRefusedException() : CoreException(ErrorCode.Conflict, ErrorDetail.ExcelNeedsReview,
    "Excel has a workbook or Protected View window that xlwright does not manage. Have a person close it, then retry.");

// Raised only for a save that clearly failed; an unknown save result must not use this type.
internal sealed class SaveFailedException(Exception? inner = null) : CoreException(ErrorCode.ExcelError,
    ErrorDetail.SaveFailed, "Saving the workbook failed and the session stays ready. Fix the cause, then retry.", inner);
