using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Spike0;

internal sealed record ProcessIdentity(int Pid, DateTime StartedAtUtc, string Name)
{
    internal static ProcessIdentity Read(Process process) =>
        new(process.Id, process.StartTime.ToUniversalTime(), process.ProcessName);

    internal bool Matches(ProcessIdentity other) => Pid == other.Pid &&
        StartedAtUtc == other.StartedAtUtc && StringComparer.OrdinalIgnoreCase.Equals(Name, other.Name);
}

internal sealed record OwnershipRecord(int Version, string Kind,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] string? WorkbookPath = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] ProcessIdentity? Process = null)
{
    internal static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        Converters = { new RoundTripUtcConverter() }
    };

    internal static OwnershipRecord Parse(string text)
    {
        try
        {
            using var document = JsonDocument.Parse(text);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object || HasDuplicateProperties(root))
                throw new FormatException();
            if (root.GetProperty("version").GetInt32() != 1)
                throw new FormatException();
            var kind = root.GetProperty("kind").GetString();
            if (kind == "completed")
            {
                if (root.EnumerateObject().Count() != 2) throw new FormatException();
                return new(1, kind);
            }
            if (kind is not ("starting" or "owned")) throw new FormatException();
            var path = root.GetProperty("workbook_path").GetString();
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path) ||
                !StringComparer.OrdinalIgnoreCase.Equals(path, Path.GetFullPath(path)))
                throw new FormatException();
            ProcessIdentity? process = null;
            if (kind == "owned")
            {
                var item = root.GetProperty("process");
                if (item.ValueKind != JsonValueKind.Object || HasDuplicateProperties(item))
                    throw new FormatException();
                var pid = item.GetProperty("pid").GetInt32();
                var name = item.GetProperty("name").GetString();
                if (pid <= 0 || string.IsNullOrWhiteSpace(name) ||
                    !DateTime.TryParseExact(item.GetProperty("started_at_utc").GetString(), "O",
                        CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var start) ||
                    start.Kind != DateTimeKind.Utc)
                    throw new FormatException();
                process = new(pid, start, name);
            }
            return new(1, kind, path, process);
        }
        catch (Exception error) when (error is JsonException or FormatException or InvalidOperationException
            or KeyNotFoundException or ArgumentException or OverflowException)
        {
            throw new SessionRecordInvalidException(error);
        }
    }

    private static bool HasDuplicateProperties(JsonElement item) =>
        item.EnumerateObject().Select(property => property.Name).Distinct(StringComparer.Ordinal).Count() !=
        item.EnumerateObject().Count();
}

internal sealed class OwnershipRecordFile : IDisposable
{
    private readonly FileStream stream;
    // Self-checks and the fixed L2 runs install this hook; the exclusive production handle stays in use.
    internal Action<string, string, FileStream>? BeforeIo { get; set; }
    internal bool CreatedNew { get; }

    private OwnershipRecordFile(FileStream stream, bool createdNew)
    {
        this.stream = stream;
        CreatedNew = createdNew;
    }

    internal static OwnershipRecordFile Acquire(string path)
    {
        try
        {
            return new(new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None), true);
        }
        catch (IOException) when (File.Exists(path))
        {
            // Open never truncates an existing record, even if validation fails.
            return new(new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None), false);
        }
    }

    internal OwnershipRecord Read()
        => Read("existing");

    private OwnershipRecord Read(string boundary)
    {
        string text = Io(boundary, boundary == "existing" ? "read" : "read-back", () =>
        {
            stream.Position = 0;
            BeforeIo?.Invoke(boundary, boundary == "existing" ? "read" : "read-back", stream);
            using var reader = new StreamReader(stream, System.Text.Encoding.UTF8,
                detectEncodingFromByteOrderMarks: false, leaveOpen: true);
            return reader.ReadToEnd();
        });
        return OwnershipRecord.Parse(text);
    }

    internal void Write(OwnershipRecord record)
    {
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(record, OwnershipRecord.JsonOptions);
        Io(record.Kind, "write", () =>
        {
            stream.Position = 0;
            BeforeIo?.Invoke(record.Kind, "write", stream);
            stream.Write(bytes);
            return true;
        });
        Io(record.Kind, "set-length", () =>
        {
            BeforeIo?.Invoke(record.Kind, "set-length", stream);
            stream.SetLength(bytes.Length);
            return true;
        });
        Io(record.Kind, "flush", () =>
        {
            BeforeIo?.Invoke(record.Kind, "flush", stream);
            stream.Flush(flushToDisk: true);
            return true;
        });
        Io(record.Kind, "read-back", () =>
        {
            if (Read(record.Kind) != record) throw new IOException("Ownership record read-back did not match.");
            return true;
        });
    }

    private static T Io<T>(string boundary, string operation, Func<T> action)
    {
        try { return action(); }
        catch (RecordIoException) { throw; }
        catch (SessionRecordInvalidException error) when (operation == "read-back")
        {
            throw new RecordIoException(boundary, operation, error);
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            throw new RecordIoException(boundary, operation, error);
        }
    }

    // Fixed L2 runs fail before any byte of the target record is written, so the previous record is retained.
    internal void InjectWriteFailure(string kind) => BeforeIo = (boundary, operation, _) =>
    {
        if (boundary == kind && operation == "write") throw new IOException("Injected L2 record write failure.");
    };

    internal void CheckPrevious()
    {
        if (CreatedNew) return;
        var record = Read();
        if (record.Kind == "completed") return;
        if (record.Kind == "starting") throw new SessionRecordInvalidException();
        Process process;
        try { process = Process.GetProcessById(record.Process!.Pid); }
        catch (ArgumentException)
        {
            Write(new(1, "completed"));
            return;
        }
        using (process)
        {
            var identity = ProcessIdentity.Read(process);
            var same = record.Process!.Matches(identity);
            throw new InvalidOperationException(same ?
                "Recorded Excel still exists. Manual recovery is required; no COM connection was attempted." :
                "Recorded PID identity does not match. Manual recovery is required.");
        }
    }

    internal void VerifyRetained(OwnershipRecord expected)
    {
        if (Read() != expected) throw new IOException("Retained ownership record changed.");
        try
        {
            using var competing = new FileStream(stream.Name, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (IOException error) when (error.HResult == unchecked((int)0x80070020)) { return; }
        throw new IOException("Expected ownership sharing violation was not observed.");
    }

    public void Dispose() => stream.Dispose();
}

internal sealed class RecordIoException(string boundary, string operation, Exception inner)
    : IOException($"SESSION_RECORD_IO_ERROR: {boundary}/{operation}", inner)
{
    internal string Boundary { get; } = boundary;
    internal string Operation { get; } = operation;
}

internal sealed class SessionRecordInvalidException(Exception? inner = null)
    : Exception("SESSION_RECORD_INVALID: preserve the record and review it manually.", inner);

// System.Text.Json's default trims fractional zeros; the adopted record requires "O".
internal sealed class RoundTripUtcConverter : JsonConverter<DateTime>
{
    public override DateTime Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
        DateTime.ParseExact(reader.GetString()!, "O", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);

    public override void Write(Utf8JsonWriter writer, DateTime value, JsonSerializerOptions options) =>
        writer.WriteStringValue(value.ToString("O", CultureInfo.InvariantCulture));
}
