using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeComparisonView : UserControl
{
    private static readonly (int Left, int Right, string Title)[] Pairs =
        [(1, 2, "取り込み1 ↔ 取り込み2"), (0, 1, "共通ベース ↔ 取り込み1"), (0, 2, "共通ベース ↔ 取り込み2")];
    private static readonly string[] Titles = ["共通ベース", "取り込み1", "取り込み2"];
    private readonly FolderMergeBook book;
    private readonly int pairCount;
    private readonly CancellationTokenSource lifetime = new();
    private readonly ComboBox pairs = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDownList,
        AccessibleName = "シート差分を比較する版" };
    private readonly Label status = Appearance.Label("");
    private readonly TextBox mergeInfo = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true,
        Height = 52, ScrollBars = ScrollBars.Vertical, AccessibleName = "選択した差分の取り込み状態と理由" };
    private readonly Panel content = new() { Dock = DockStyle.Fill };
    private readonly Dictionary<int, IReadOnlyList<SheetDiff>> comparisons = [];
    private PreviewOptions options = new();
    private Workbook?[]? versions;
    private WorkbookView? current;
    private bool loading;
    private string loadMessage = "";

    internal MergeComparisonView(FolderMergeBook book, bool useBase)
    {
        this.book = book;
        pairCount = useBase ? 3 : 1;
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        pairs.Items.AddRange(Pairs.Take(pairCount).Select(pair => pair.Title).ToArray());
        pairs.SelectedIndex = 0;
        pairs.Margin = new Padding(6, 7, 0, 4);
        pairs.SelectedIndexChanged += (_, _) => { if (!loading) ShowPair(); };
        toolbar.Controls.AddRange([pairs, status]);
        Appearance.Row(layout, toolbar);
        Appearance.Row(layout, mergeInfo);
        Appearance.Row(layout, content, fill: true);
        Controls.Add(layout);
    }

    internal event Action<SheetDiff?, Difference?, MergeCellAddress?>? DifferenceSelected;

    internal void SetOptions(PreviewOptions next)
    {
        options = next;
        current?.SetPreviewOptions(next);
    }

    internal void ShowMergeInfo(string text) => mergeInfo.Text = text;

    internal async void LoadComparison()
    {
        if (IsDisposed || versions is not null || loading) return;
        loading = true;
        pairs.Enabled = false;
        status.Text = "シート差分を読み込んでいます…";
        var token = lifetime.Token;
        try
        {
            var result = await Task.Run(() =>
            {
                var loaded = book.Plan?.ReadPreview(token) ?? new Workbook?[3];
                var messages = new List<string>();
                var paths = new[] { book.BasePath, book.ProductionPath, book.IncomingPath };
                for (var i = 0; i < paths.Length; i++)
                {
                    if (book.Plan is not null || paths[i] is not string path) continue;
                    try { loaded[i] = XlsxReader.Read(path, token); }
                    catch (Exception error) when (FolderComparer.IsReadError(error))
                    {
                        messages.Add($"{Titles[i]}を読み込めません。{error.Message}");
                    }
                }
                var differences = new Dictionary<int, IReadOnlyList<SheetDiff>>();
                for (var i = 0; i < pairCount; i++)
                {
                    var (left, right) = PairBooks(loaded, i);
                    differences[i] = WorkbookComparer.Compare(left, right, token: token);
                }
                return (Versions: loaded, Differences: differences, Message: string.Join(" ", messages));
            }, token);
            if (IsDisposed || token.IsCancellationRequested) return;
            versions = result.Versions;
            loadMessage = result.Message;
            foreach (var (index, differences) in result.Differences)
            {
                comparisons[index] = differences;
                pairs.Items[index] = $"{Pairs[index].Title}（{differences.Sum(sheet => sheet.Differences.Count):N0} 件）";
            }
            if (comparisons[pairs.SelectedIndex].All(sheet => sheet.Differences.Count == 0))
                pairs.SelectedIndex = Enumerable.Range(0, pairCount)
                    .FirstOrDefault(index => comparisons[index].Any(sheet => sheet.Differences.Count > 0));
            ShowPair();
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error))
        {
            if (!IsDisposed) status.Text = error.Message;
        }
        finally
        {
            loading = false;
            if (!IsDisposed) pairs.Enabled = true;
        }
    }

    private static (Workbook Left, Workbook Right) PairBooks(Workbook?[] books, int index)
    {
        var pair = Pairs[index];
        var left = books[pair.Left] ?? new Workbook([], books[pair.Right]?.Date1904 ?? false);
        return (left, books[pair.Right] ?? new Workbook([], left.Date1904));
    }

    private void ShowPair()
    {
        if (IsDisposed || Disposing || versions is not { } books) return;
        current?.Dispose();
        current = null;
        DifferenceSelected?.Invoke(null, null, null);
        var index = pairs.SelectedIndex;
        var pair = Pairs[index];
        var (left, right) = PairBooks(books, index);
        var differences = comparisons[index];
        var paths = new[] { book.BasePath, book.ProductionPath, book.IncomingPath };
        var view = new WorkbookView(paths[pair.Left] ?? "", paths[pair.Right] ?? "", left, right, differences)
            { Dock = DockStyle.Fill };
        current = view;
        view.SetPreviewOptions(options);
        view.DifferenceSelected += (sheet, difference) =>
        {
            var row = pair.Left == 1 ? difference?.LeftRow : difference?.RightRow ?? difference?.LeftRow;
            var column = difference?.Column;
            var cell = sheet is not null && row is int r && column is int c
                ? new MergeCellAddress(sheet.Name, r, c) : null;
            DifferenceSelected?.Invoke(sheet, difference, cell);
        };
        content.Controls.Add(view);
        view.ShowDifferenceList();
        status.Text = loadMessage.Length > 0 ? loadMessage : $"{differences.Count} シート　差分 {differences.Sum(sheet => sheet.Differences.Count):N0} 件";
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && !IsDisposed)
        {
            lifetime.Cancel();
            lifetime.Dispose();
        }
        base.Dispose(disposing);
    }
}
