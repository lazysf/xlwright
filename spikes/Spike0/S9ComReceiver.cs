using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using System.Windows.Forms;

namespace Spike0;

internal sealed record S9ComReady(string Nonce, string Scenario, int Pid, string Build, string Reference);
internal sealed record S9ComReceived(string Nonce, string Scenario, int Thread, int BodyCalls, int Incoming, int Rejected, bool Restored);

internal static class S9ComReceiver
{
    internal static readonly string[] Cases = ["accept", "retry", "reject", "limit", "deadline", "failure", "late"];
    internal static string Build => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
        File.ReadAllBytes(typeof(S9ComReceiver).Assembly.Location)));

    internal static Task<int> Run(string scenario, string nonce)
    {
        if (!Cases.Contains(scenario) || !Guid.TryParseExact(nonce, "N", out _))
            throw new ArgumentException("Invalid fixed receiver request.");
        var completion = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try { completion.SetResult(Serve(scenario, nonce)); }
            catch (Exception error) { completion.SetException(error); }
        }) { IsBackground = true, Name = "S9 controlled COM receiver" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private static int Serve(string scenario, string nonce)
    {
        using var control = new Control();
        _ = control.Handle;
        var target = new S9MemoryStream(scenario);
        using var filter = new S9ReceiverFilter(scenario);
        filter.Register();
        string reference = S9ComMarshalling.Export(target);
        Console.WriteLine(JsonSerializer.Serialize(new S9ComReady(nonce, scenario, Environment.ProcessId, Build, reference)));
        // Only the parent's dedicated stdin controls normal receiver shutdown.
        _ = Task.Run(async () =>
        {
            try
            {
                string? command = await Console.In.ReadLineAsync();
                if (command is not null and not "stop")
                    Console.Error.WriteLine("Unexpected receiver control command.");
            }
            finally { control.BeginInvoke((Action)Application.ExitThread); }
        });
        Application.Run();
        filter.Dispose();
        Console.WriteLine(JsonSerializer.Serialize(new S9ComReceived(nonce, scenario,
            Environment.CurrentManagedThreadId, target.BodyCalls, filter.Incoming, filter.Rejected, filter.Restored)));
        GC.KeepAlive(target);
        return 0;
    }
}

// Standard OLE marshaling keeps the fixture on its receiver STA; no type library or registration.
[ComVisible(true), ClassInterface(ClassInterfaceType.None)]
public sealed class S9MemoryStream : StandardOleMarshalObject, IStream
{
    private readonly string scenario;
    private readonly int owner = Environment.CurrentManagedThreadId;
    internal int BodyCalls { get; private set; }
    internal S9MemoryStream(string scenario) => this.scenario = scenario;
    public void Commit(int flags)
    {
        if (Environment.CurrentManagedThreadId != owner) throw new COMException("Receiver STA mismatch.");
        BodyCalls++;
        if (scenario == "late") Thread.Sleep(500);
        if (scenario == "failure") throw new COMException("Accepted fixture failure.", unchecked((int)0x80004005));
    }
    // The fixed probe uses only Commit; other stream operations are deliberately unsupported.
    public void Read(byte[] buffer, int count, nint read) => throw new NotSupportedException();
    public void Write(byte[] buffer, int count, nint written) => throw new NotSupportedException();
    public void Seek(long offset, int origin, nint position) => throw new NotSupportedException();
    public void SetSize(long size) => throw new NotSupportedException();
    public void CopyTo(IStream stream, long count, nint read, nint written) => throw new NotSupportedException();
    public void Revert() => throw new NotSupportedException();
    public void LockRegion(long offset, long count, int type) => throw new NotSupportedException();
    public void UnlockRegion(long offset, long count, int type) => throw new NotSupportedException();
    public void Stat(out STATSTG stat, int flags) => throw new NotSupportedException();
    public void Clone(out IStream stream) => throw new NotSupportedException();
}

[ComVisible(true), ClassInterface(ClassInterfaceType.None)]
public sealed class S9ReceiverFilter : IS9MessageFilter, IDisposable
{
    private readonly string scenario;
    private IS9MessageFilter? previous;
    private bool installed;
    internal int Incoming { get; private set; }
    internal int Rejected { get; private set; }
    internal bool Restored { get; private set; }
    internal S9ReceiverFilter(string scenario) => this.scenario = scenario;
    internal void Register()
    {
        if (CoRegisterMessageFilter(this, out previous) != 0) throw new InvalidOperationException("Receiver filter registration failed.");
        installed = true;
    }
    public int HandleInComingCall(int callType, nint caller, int ticks, nint info)
    {
        if (info == 0) return 0;
        var method = Marshal.PtrToStructure<InterfaceInfo>(info);
        // Do not reject setup, QueryInterface, or lifetime calls.
        if (method.Iid != typeof(IStream).GUID || method.Method != 8) return 0;
        Incoming++;
        int rejection = scenario switch
        {
            "retry" when Rejected < 2 => 2,
            "reject" => 1,
            "limit" or "deadline" => 2,
            _ => 0
        };
        if (rejection != 0) Rejected++;
        return rejection;
    }
    public int RetryRejectedCall(nint callee, int ticks, int type) => -1;
    public int MessagePending(nint callee, int ticks, int type) => 2;
    public void Dispose()
    {
        if (!installed) return;
        if (CoRegisterMessageFilter(previous, out _) != 0) throw new InvalidOperationException("Receiver filter restore failed.");
        installed = false;
        Restored = true;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct InterfaceInfo { internal nint Unknown; internal Guid Iid; internal ushort Method; }
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IS9MessageFilter? filter, out IS9MessageFilter? previous);
}

internal static class S9ComMarshalling
{
    internal static string Export(object target)
    {
        Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(0, true, out var stream));
        try
        {
            Guid iid = typeof(IStream).GUID;
            Marshal.ThrowExceptionForHR(CoMarshalInterface(stream, ref iid, target, 0, 0, 0));
            stream.Stat(out var stat, 1);
            if (stat.cbSize is <= 0 or > 65536) throw new InvalidDataException("Unexpected marshal size.");
            byte[] bytes = new byte[(int)stat.cbSize];
            stream.Seek(0, 0, 0);
            stream.Read(bytes, bytes.Length, 0);
            return Convert.ToBase64String(bytes);
        }
        finally { Marshal.FinalReleaseComObject(stream); }
    }
    internal static IStream Import(string reference)
    {
        if (reference.Length > 90000) throw new InvalidDataException("Marshal reference too large.");
        byte[] bytes = Convert.FromBase64String(reference);
        Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(0, true, out var stream));
        try
        {
            stream.Write(bytes, bytes.Length, 0);
            stream.Seek(0, 0, 0);
            Guid iid = typeof(IStream).GUID;
            Marshal.ThrowExceptionForHR(CoUnmarshalInterface(stream, ref iid, out nint pointer));
            try { return (IStream)Marshal.GetObjectForIUnknown(pointer); }
            finally { Marshal.Release(pointer); }
        }
        finally { Marshal.FinalReleaseComObject(stream); }
    }
    [DllImport("ole32.dll")]
    private static extern int CreateStreamOnHGlobal(nint memory, [MarshalAs(UnmanagedType.Bool)] bool delete, out IStream stream);
    [DllImport("ole32.dll")]
    private static extern int CoMarshalInterface(IStream stream, ref Guid iid,
        [MarshalAs(UnmanagedType.IUnknown)] object target, int context, nint reserved, int flags);
    [DllImport("ole32.dll")]
    private static extern int CoUnmarshalInterface(IStream stream, ref Guid iid, out nint pointer);
}
