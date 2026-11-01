namespace XlWright.Core;

internal static class RecordTransition
{
    internal static void Run(ExecutionGate gate, OwnershipRecordFile file, OwnershipRecord record, Action next)
    {
        gate.Step("record-" + record.Kind, () => file.Write(record));
        next();
    }
}
