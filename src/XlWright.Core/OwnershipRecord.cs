using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace XlWright.Core;

/// <summary>The process of an owned record, matched by all three facts so a reused PID is not mistaken for it.</summary>
public sealed record RecordedProcess(int Pid, DateTime StartedAtUtc, string Name);

/// <summary>The ownership record in session.json (docs/architecture.md "所有記録の形式").</summary>
public sealed record OwnershipRecord(string State, DateTime MarkedAtUtc, string Workbook, RecordedProcess? Process)
{
    public const string Starting = "starting";
    public const string Owned = "owned";
    public const string Completed = "completed";

    /// <summary>
    /// Reads strict JSON: no comments, trailing commas, duplicate or unknown keys, and process only when owned.
    /// Returns null for anything else, so the start stops instead of guessing.
    /// </summary>
    public static OwnershipRecord? Parse(string text)
    {
        try
        {
            using var document = JsonDocument.Parse(text);
            var root = Properties(document.RootElement, "version", "state", "marked_at_utc", "workbook", "process");
            if (root is null
                || !root.TryGetValue("version", out var version) || version.ValueKind != JsonValueKind.Number || version.GetRawText() != "1"
                || !root.TryGetValue("state", out var state) || state.ValueKind != JsonValueKind.String || state.GetString() is not (Starting or Owned or Completed)
                || !root.TryGetValue("marked_at_utc", out var marked) || Time(marked) is not { } markedAt
                || !root.TryGetValue("workbook", out var workbook) || workbook.ValueKind != JsonValueKind.String)
            {
                return null;
            }

            var hasProcess = root.TryGetValue("process", out var process);
            if (hasProcess != (state.GetString() == Owned))
            {
                return null;
            }

            RecordedProcess? recorded = null;
            if (hasProcess)
            {
                var facts = Properties(process, "pid", "started_at_utc", "name");
                if (facts is null || facts.Count != 3
                    || facts["pid"].ValueKind != JsonValueKind.Number || !facts["pid"].TryGetInt32(out var pid) || pid <= 0
                    || Time(facts["started_at_utc"]) is not { } started
                    || facts["name"].ValueKind != JsonValueKind.String)
                {
                    return null;
                }

                recorded = new RecordedProcess(pid, started, facts["name"].GetString()!);
            }

            return new OwnershipRecord(state.GetString()!, markedAt, workbook.GetString()!, recorded);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    public string ToJson()
    {
        var json = new JsonObject
        {
            ["version"] = 1,
            ["state"] = State,
            ["marked_at_utc"] = MarkedAtUtc.ToString("O", CultureInfo.InvariantCulture),
            ["workbook"] = Workbook,
        };
        if (Process is not null)
        {
            json["process"] = new JsonObject
            {
                ["pid"] = Process.Pid,
                ["started_at_utc"] = Process.StartedAtUtc.ToString("O", CultureInfo.InvariantCulture),
                ["name"] = Process.Name,
            };
        }

        return json.ToJsonString(JsonText.Options);
    }

    // The properties of an object, or null when it is not an object or has a duplicate or unknown key.
    private static Dictionary<string, JsonElement>? Properties(JsonElement element, params string[] known)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        var properties = new Dictionary<string, JsonElement>();
        foreach (var property in element.EnumerateObject())
        {
            if (!known.Contains(property.Name) || !properties.TryAdd(property.Name, property.Value))
            {
                return null;
            }
        }

        return properties;
    }

    // A round-trip UTC time, such as 2026-11-02T01:23:45.6789012Z.
    private static DateTime? Time(JsonElement element) =>
        element.ValueKind == JsonValueKind.String
        && DateTime.TryParseExact(element.GetString(), "O", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var time)
        && time.Kind == DateTimeKind.Utc
            ? time
            : null;
}
