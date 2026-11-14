namespace ExcelDiff.Desktop.Views;

internal static class RecentTargets
{
    private static string SettingsPath => LocalSettings.PathFor("last-comparison.json");

    internal static Targets? Load() => Load(SettingsPath);

    internal static Targets? Load(string path)
    {
        var targets = LocalSettings.Load<Targets>(path);
        return targets is not null && !string.IsNullOrWhiteSpace(targets.Left) &&
            !string.IsNullOrWhiteSpace(targets.Right) ? targets : null;
    }

    internal static void Save(Targets targets) => Save(SettingsPath, targets);

    internal static void Save(string path, Targets targets) => LocalSettings.Save(path, targets);
}
