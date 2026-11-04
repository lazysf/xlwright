using System.Text;

namespace XlWright.Core;

/// <summary>
/// The app lock: &lt;app&gt;/.xlwright/session.json held open with FileShare.None for the whole session, and the
/// ownership record written through the same handle (docs/architecture.md "実行時の状態と所有記録").
/// When the server is killed, the OS closes the handle and frees the app.
/// </summary>
public sealed class AppLock : IDisposable
{
    private const int SharingViolation = unchecked((int)0x80070020);
    private const int FileExists = unchecked((int)0x80070050);

    private static readonly UTF8Encoding Strict = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private readonly FileStream _file;

    private AppLock(FileStream file, bool created)
    {
        _file = file;
        Created = created;
    }

    /// <summary>True when this start created the file; only then is it the first start of the app.</summary>
    public bool Created { get; }

    public static string RelativePath(AppFolder app) => $"{app.RelativePath}/.xlwright/session.json";

    /// <summary>Takes the lock, or returns null when another server holds it. Other I/O failures throw.</summary>
    public static AppLock? Take(AppFolder app)
    {
        var folder = Path.Combine(app.FullPath, ".xlwright");
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "session.json");
        try
        {
            return new AppLock(new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None), created: true);
        }
        catch (IOException e) when (e.HResult == FileExists)
        {
        }

        try
        {
            return new AppLock(new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None), created: false);
        }
        catch (IOException e) when (e.HResult == SharingViolation)
        {
            return null;
        }
    }

    /// <summary>
    /// Takes the lock on the record only when it exists, never creating it (docs/api.md "表示要求のコマンド").
    /// Returns null when there is no record (<paramref name="missing"/> is true) or another server holds it.
    /// </summary>
    public static AppLock? TakeExisting(AppFolder app, out bool missing)
    {
        var path = Path.Combine(app.FullPath, ".xlwright", "session.json");
        missing = false;
        try
        {
            return new AppLock(new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None), created: false);
        }
        catch (Exception e) when (e is FileNotFoundException or DirectoryNotFoundException)
        {
            missing = true;
            return null;
        }
        catch (IOException e) when (e.HResult == SharingViolation)
        {
            return null;
        }
    }

    /// <summary>The whole file as UTF-8 without a BOM. Throws when it cannot be read or decoded.</summary>
    public string Read()
    {
        _file.Position = 0;
        var bytes = new byte[_file.Length];
        _file.ReadExactly(bytes);
        return Strict.GetString(bytes);
    }

    /// <summary>
    /// Writes from the start without truncating first, sets the length, flushes to disk, and reads it back
    /// (docs/architecture.md "所有記録の形式"). Throws when the read-back differs.
    /// </summary>
    public void Write(OwnershipRecord record)
    {
        var text = record.ToJson();
        var bytes = Strict.GetBytes(text);
        _file.Position = 0;
        _file.Write(bytes);
        _file.SetLength(bytes.Length);
        _file.Flush(flushToDisk: true);
        if (Read() != text)
        {
            throw new IOException("The ownership record read back differs from what was written.");
        }
    }

    public void Dispose() => _file.Dispose();
}
