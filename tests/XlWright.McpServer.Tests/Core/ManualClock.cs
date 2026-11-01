namespace XlWright.McpServer.Tests.Core;

// Deadlines are decided by the gate's clock, so tests move time explicitly instead of sleeping.
internal sealed class ManualClock : TimeProvider
{
    private long timestamp;
    public override long TimestampFrequency => TimeSpan.TicksPerSecond;
    public override long GetTimestamp() => Interlocked.Read(ref timestamp);
    internal void Advance(TimeSpan duration) => Interlocked.Add(ref timestamp, duration.Ticks);
}
