using System.Diagnostics;

namespace XlWright.Core.Com;

/// <summary>Finds processes through the same handle reads as <see cref="WindowsProcess"/>.</summary>
public sealed class WindowsProcessQuery : IProcessQuery
{
    // A read failure is never taken for an ended process (docs/architecture.md "所有記録の形式").
    public ProcessLookup Find(RecordedProcess recorded)
    {
        var process = WindowsProcess.Open(recorded.Pid);
        var presence = Presence(process, recorded);
        if (presence == ProcessPresence.Running)
        {
            return new ProcessLookup(presence, process);
        }

        process.Dispose();
        return new ProcessLookup(presence);
    }

    private static ProcessPresence Presence(WindowsProcess process, RecordedProcess recorded)
    {
        if (process.OpenError == WindowsProcess.NoSuchProcess)
        {
            return ProcessPresence.Gone;
        }

        var facts = process.Facts;
        if (process.OpenError != 0 || facts.CreatedUtc == DateTime.MinValue)
        {
            return ProcessPresence.Unknown;
        }

        // Another creation time is a reused PID.
        if (facts.CreatedUtc != recorded.StartedAtUtc || process.WaitForExit(TimeSpan.Zero))
        {
            return ProcessPresence.Gone;
        }

        return string.Equals(Path.GetFileNameWithoutExtension(facts.ImageName), recorded.Name, StringComparison.OrdinalIgnoreCase)
            ? ProcessPresence.Running
            : ProcessPresence.Unknown;
    }

    // A process whose creation time or command line cannot be read is counted, so an unknown Excel stops the start.
    public IReadOnlyList<ProcessFacts> AutomationExcelsSince(DateTime utc)
    {
        var found = new List<ProcessFacts>();
        foreach (var listed in Process.GetProcessesByName("EXCEL"))
        {
            using (listed)
            using (var process = WindowsProcess.Open(listed.Id))
            {
                var facts = process.Facts;
                var unreadable = facts.CreatedUtc == DateTime.MinValue || facts.CommandLine is null;
                if (unreadable || (facts.CreatedUtc > utc && facts.CommandLine!.Contains("/automation", StringComparison.OrdinalIgnoreCase)))
                {
                    found.Add(facts);
                }
            }
        }

        return found;
    }
}
