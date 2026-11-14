using System.Security;
using System.Text.Json;

namespace ExcelDiff.Desktop.Views;

internal static class LocalSettings
{
    internal static string PathFor(string name) => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelDiff", name);

    internal static T? Load<T>(string path) where T : class
    {
        try
        {
            using var file = File.OpenRead(path);
            return JsonSerializer.Deserialize<T>(file);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException or JsonException)
        {
            return null;
        }
    }

    internal static void Save<T>(string path, T value)
    {
        var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            using (var file = File.Create(temporary))
                JsonSerializer.Serialize(file, value);
            File.Move(temporary, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or SecurityException)
        {
            try { File.Delete(temporary); }
            catch (Exception cleanup) when (cleanup is IOException or UnauthorizedAccessException or SecurityException) { }
        }
    }
}
