using System.Security;
using System.Text.Json;

namespace ExcelDiff.Desktop.Views;

internal static class RecentTargets
{
    private static string SettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelDiff", "last-comparison.json");

    internal static Targets? Load() => Load(SettingsPath);

    internal static Targets? Load(string path)
    {
        try
        {
            using var file = File.OpenRead(path);
            var targets = JsonSerializer.Deserialize<Targets>(file);
            return targets is not null && !string.IsNullOrWhiteSpace(targets.Left) &&
                !string.IsNullOrWhiteSpace(targets.Right) ? targets : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException or JsonException)
        {
            return null;
        }
    }

    internal static void Save(Targets targets) => Save(SettingsPath, targets);

    internal static void Save(string path, Targets targets)
    {
        var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            using (var file = File.Create(temporary))
                JsonSerializer.Serialize(file, targets);
            File.Move(temporary, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException)
        {
            try { File.Delete(temporary); }
            catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException or SecurityException) { }
        }
    }
}
