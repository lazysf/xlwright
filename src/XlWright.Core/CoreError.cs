namespace XlWright.Core;

// The six public codes in docs/api.md. Do not add codes; add a detail instead.
internal enum ErrorCode { InvalidRequest, NotFound, Conflict, ExcelError, StateUnknown, InternalError }

internal static class ErrorDetail
{
    internal const string SessionRecordInvalid = "SESSION_RECORD_INVALID";
    internal const string SessionRecordIoError = "SESSION_RECORD_IO_ERROR";
    internal const string ExcelNeedsReview = "EXCEL_NEEDS_REVIEW";
    internal const string ProcessQueryFailed = "PROCESS_QUERY_FAILED";
    internal const string SaveFailed = "SAVE_FAILED";
}

// Core decides the public code and detail where the failure is known, so the MCP layer only formats them.
internal class CoreException(ErrorCode code, string detail, string message, Exception? inner = null)
    : Exception(message, inner)
{
    internal ErrorCode Code { get; } = code;
    internal string Detail { get; } = detail;
}
