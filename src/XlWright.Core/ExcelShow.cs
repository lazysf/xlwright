namespace XlWright.Core;

/// <summary>Reads and shows the workbook windows of a process (docs/api.md "表示要求のコマンド"). It never touches COM.</summary>
public interface IWorkbookWindows
{
    /// <summary>Shows each XLMAIN of the process that holds an EXCEL7 and brings it to the front; returns how many.</summary>
    int Show(int pid);

    /// <summary>Whether the process has an EXCEL7 window under an XLMAIN.</summary>
    bool HasWorkbookWindow(int pid);
}

/// <summary>
/// The show command (docs/api.md "表示要求のコマンド"): holds the app lock, shows the recorded dedicated Excel to the
/// person, and after they type done, ends it when it is left without a workbook window. The record is never
/// written; the next start completes it. Returns the exit code: 0 nothing left, 1 stopped.
/// </summary>
public sealed class ExcelShow(IProcessQuery processes, IWorkbookWindows windows, TextReader input, TextWriter output, TimeSpan terminateExit)
{
    public int Run(AppFolder app)
    {
        var record = AppLock.RelativePath(app);
        using var appLock = AppLock.TakeExisting(app, out var missing);
        if (missing)
        {
            output.WriteLine($"{record} does not exist, so no dedicated Excel of this app is left.");
            return 0;
        }

        if (appLock is null)
        {
            output.WriteLine($"{app.RelativePath} is in use by another conversation. End its session or close it, then run this again.");
            return 1;
        }

        OwnershipRecord? read;
        try
        {
            read = OwnershipRecord.Parse(appLock.Read());
        }
        catch (Exception e) when (e is IOException or System.Text.DecoderFallbackException)
        {
            read = null;
        }

        switch (read?.State)
        {
            case null:
                output.WriteLine($"{record} is empty or not in the expected form. Read it, check that the Excel it names is not running, then delete the file.");
                return 1;
            case OwnershipRecord.Completed:
                output.WriteLine("No dedicated Excel of this app is left.");
                return 0;
            case OwnershipRecord.Starting:
                var started = processes.AutomationExcelsSince(read.MarkedAtUtc);
                output.WriteLine($"An earlier start stopped while starting Excel, so no Excel is recorded. Excel processes started from COM since {read.MarkedAtUtc:O}: " +
                    (started.Count == 0 ? "none." : string.Join(", ", started.Select(facts => $"PID {facts.Pid} created {facts.CreatedUtc:O}")) + "."));
                output.WriteLine("Check who owns each one in Task Manager before ending it.");
                return 1;
        }

        return ShowOwned(read.Process!);
    }

    private int ShowOwned(RecordedProcess recorded)
    {
        var lookup = processes.Find(recorded);
        if (lookup.Presence == ProcessPresence.Gone)
        {
            output.WriteLine($"The recorded dedicated Excel (PID {recorded.Pid}) is not running. The next start completes the record.");
            return 0;
        }

        if (lookup.Presence == ProcessPresence.Unknown)
        {
            output.WriteLine($"The recorded dedicated Excel (PID {recorded.Pid}) could not be checked, so it was left untouched. Check it in Task Manager.");
            return 1;
        }

        using var process = lookup.Process!;
        if (process.Facts.CommandLine is not { } commandLine || !commandLine.Contains("/automation", StringComparison.OrdinalIgnoreCase))
        {
            output.WriteLine($"PID {recorded.Pid} could not be confirmed as started from COM, so it was left untouched. Check it in Task Manager.");
            return 1;
        }

        output.WriteLine(windows.Show(recorded.Pid) > 0
            ? $"The dedicated Excel (PID {recorded.Pid}) is shown. Check it, save what you need, and close it."
            : $"The dedicated Excel (PID {recorded.Pid}) has no workbook window to show.");
        while (true)
        {
            output.WriteLine("Type done and press Enter when you have finished.");
            var line = input.ReadLine();
            if (line is null)
            {
                output.WriteLine("Input closed, so the Excel was left as it is.");
                return 1;
            }

            if (!string.Equals(line.Trim(), "done", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (process.WaitForExit(TimeSpan.Zero))
            {
                output.WriteLine("The dedicated Excel has ended.");
                return 0;
            }

            if (windows.HasWorkbookWindow(recorded.Pid))
            {
                output.WriteLine("A workbook window is still open. Close it, then type done again.");
                continue;
            }

            // Locked, identified with /automation, and without a workbook window, as in the cleanup (人の判断, 2026-11-04).
            process.Terminate();
            if (process.WaitForExit(terminateExit))
            {
                output.WriteLine("The dedicated Excel was left running without a workbook window, so it was ended.");
                return 0;
            }

            output.WriteLine($"The dedicated Excel (PID {recorded.Pid}) was asked to end but its end could not be confirmed. Check it in Task Manager.");
            return 1;
        }
    }
}
