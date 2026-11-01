namespace XlWright.Core;

// The six public codes in docs/api.md. Do not add codes; add a detail instead.
internal enum ErrorCode { InvalidRequest, NotFound, Conflict, ExcelError, StateUnknown, InternalError }

internal static class ErrorDetail
{
    internal const string InvalidArgument = "INVALID_ARGUMENT";
    internal const string InvalidAddress = "INVALID_ADDRESS";
    internal const string ShapeMismatch = "SHAPE_MISMATCH";
    internal const string InvalidConfig = "INVALID_CONFIG";
    internal const string NotImplemented = "NOT_IMPLEMENTED";
    internal const string ConfigNotFound = "CONFIG_NOT_FOUND";
    internal const string PathNotFound = "PATH_NOT_FOUND";
    internal const string SheetNotFound = "SHEET_NOT_FOUND";
    internal const string OperationBusy = "OPERATION_BUSY";
    internal const string SessionActive = "SESSION_ACTIVE";
    internal const string SessionNotStarted = "SESSION_NOT_STARTED";
    internal const string SessionIsolated = "SESSION_ISOLATED";
    internal const string ProjectLocked = "PROJECT_LOCKED";
    internal const string WorkbookLocked = "WORKBOOK_LOCKED";
    internal const string WorkbookReadOnly = "WORKBOOK_READ_ONLY";
    internal const string SessionRecordInvalid = "SESSION_RECORD_INVALID";
    internal const string SessionRecordIoError = "SESSION_RECORD_IO_ERROR";
    internal const string ExcelNeedsReview = "EXCEL_NEEDS_REVIEW";
    internal const string ProcessQueryFailed = "PROCESS_QUERY_FAILED";
    internal const string ExcelStartFailed = "EXCEL_START_FAILED";
    internal const string WorkbookOpenFailed = "WORKBOOK_OPEN_FAILED";
    internal const string SaveFailed = "SAVE_FAILED";
    internal const string MergedCell = "MERGED_CELL";
    internal const string CellProtected = "CELL_PROTECTED";
    internal const string ExcelRejected = "EXCEL_REJECTED";
    internal const string WriteNotVerified = "WRITE_NOT_VERIFIED";
    internal const string OperationTimeout = "OPERATION_TIMEOUT";
    internal const string ExcelDisconnected = "EXCEL_DISCONNECTED";
    internal const string InternalFailure = "INTERNAL_FAILURE";
}

// Core decides the public code and detail where the failure is known, so the MCP layer only formats them.
internal class CoreException(ErrorCode code, string detail, string message, Exception? inner = null)
    : Exception(message, inner)
{
    internal ErrorCode Code { get; } = code;
    internal string Detail { get; } = detail;

    internal string CodeName => Code switch
    {
        ErrorCode.InvalidRequest => "INVALID_REQUEST",
        ErrorCode.NotFound => "NOT_FOUND",
        ErrorCode.Conflict => "CONFLICT",
        ErrorCode.ExcelError => "EXCEL_ERROR",
        ErrorCode.StateUnknown => "STATE_UNKNOWN",
        _ => "INTERNAL_ERROR"
    };
}
