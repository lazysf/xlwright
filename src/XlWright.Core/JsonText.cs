using System.Text.Encodings.Web;
using System.Text.Json;

namespace XlWright.Core;

/// <summary>
/// The JSON options for response bodies and result files. Text other than ASCII is written as is, so a Japanese
/// character counts as one toward the size limits (docs/notes/response-size.md). The output is never put into HTML.
/// </summary>
public static class JsonText
{
    public static JsonSerializerOptions Options { get; } = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
}
