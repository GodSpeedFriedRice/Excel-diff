using ExcelDiff.Desktop.Views;

namespace ExcelDiff.Desktop;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        using var window = new CompareWindow();
        Application.Run(window);
    }
}
