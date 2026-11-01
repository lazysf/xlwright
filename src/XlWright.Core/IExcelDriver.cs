namespace XlWright.Core;

// The COM boundary. Every member runs on the session's STA and returns plain values, never COM references.
// ExcelSession decides the order and the outcome, which lets tests exercise it without Excel.
internal interface IExcelDriver
{
    void Generate();
    ProcessIdentity Identify();
    void Open(string workbookPath);
    CellContent Read(string sheet, string address, bool values, bool formulas);
    void Write(string sheet, string address, CellWrite write);
    void Save();
    void VerifyEndSafety(string workbookPath, bool requireEmpty);
    void Close();
    void Quit();
    bool ReleaseAndWaitForExit(TimeSpan timeout);
}
