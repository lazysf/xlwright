using System.Reflection;
using System.Text.Json.Serialization;

namespace XlWright.Core;

/// <summary>
/// The name, version, and build id that server_info returns.
/// </summary>
public sealed record ServerIdentity(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("version")] string Version,
    [property: JsonPropertyName("build_id")] string BuildId)
{
    public const string ServerName = "xlwright";

    public static ServerIdentity Current { get; } = FromAssembly(typeof(ServerIdentity).Assembly);

    // Version comes from <Version> in Directory.Build.props; the build id from -p:XlWrightBuildId.
    private static ServerIdentity FromAssembly(Assembly assembly)
    {
        var version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? throw new InvalidOperationException("The assembly has no informational version.");
        var buildId = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()
            .SingleOrDefault(attribute => attribute.Key == "XlWrightBuildId")?.Value
            ?? throw new InvalidOperationException("The assembly has no XlWrightBuildId metadata.");
        return new ServerIdentity(ServerName, version, buildId);
    }
}
