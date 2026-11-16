using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergePreview : UserControl
{
    private static readonly string[] Titles = ["共通ベース", "取り込み1", "取り込み2"];
    private readonly FolderMergeBook book;
    private readonly bool useBase;
    private readonly CancellationTokenSource lifetime = new();
    private readonly ComboBox sheets = new() { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "プレビューするシート" };
    private readonly ComboBox pairs = new() { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "プレビューする版" };
    private readonly Label status = Appearance.Label("");
    private readonly TextBox details = new()
    {
        Dock = DockStyle.Top, Height = 64, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
        AccessibleName = "選択したセルの差分"
    };
    private readonly TableLayoutPanel split = new() { ColumnCount = 3, RowCount = 1, Dock = DockStyle.Fill };
    private readonly FullSheetPane[] panes = [new(), new(), new()];
    private readonly Control[] panels = new Control[3];
    private readonly Label[] headings = new Label[3];
    private readonly WheelZoom zoom = new();
    private readonly Dictionary<(string Sheet, int Left, int Right), IReadOnlyList<Difference>> comparisons = [];
    private readonly PaneDifferences[] displayed = new PaneDifferences[3];
    private PreviewOptions options = new();
    private Workbook?[]? versions;
    private MergeCellAddress? pendingReveal;
    private MergeCellAddress? selection;
    private string loadMessage = "";
    private (string Sheet, int Hidden)? markedScope;
    private CancellationTokenSource? comparing;
    private int hiddenPane;
    private bool loading;
    private bool selecting;

    internal MergePreview(FolderMergeBook book, bool useBase)
    {
        this.book = book;
        this.useBase = useBase;
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var commands = Appearance.Toolbar();
        sheets.Margin = pairs.Margin = new Padding(6, 7, 0, 4);
        pairs.Items.AddRange(["取り込み1 ↔ 取り込み2", "共通ベース ↔ 取り込み1", "共通ベース ↔ 取り込み2"]);
        pairs.SelectedIndex = 0;
        pairs.SelectedIndexChanged += (_, _) => SetColumns();
        sheets.SelectedIndexChanged += (_, _) =>
        {
            if (selecting) return;
            selection = pendingReveal = null;
            DrawSheet();
        };
        commands.Controls.AddRange([sheets, pairs, status]);
        Appearance.Row(layout, commands);
        for (var i = 0; i < panes.Length; i++)
        {
            var index = i;
            var pane = panes[i];
            var panel = Appearance.Stack();
            var heading = Appearance.Label(Titles[i]);
            panel.Margin = new Padding(1, 0, 0, 0);
            Appearance.Row(panel, heading);
            Appearance.Row(panel, pane, fill: true);
            headings[i] = heading;
            panels[i] = panel;
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100 / 3f));
            split.Controls.Add(panel, i, 0);
            pane.Scrolled += (_, _) =>
            {
                foreach (var other in panes)
                    if (other != pane) other.ScrollTo(pane.TopRow, pane.LeftColumn);
            };
            pane.ZoomRequested += delta =>
            {
                if (!zoom.Apply(delta)) return;
                foreach (var other in panes) other.SetZoom(zoom.Scale);
            };
            pane.Picked += pick => ShowPicked(index, pick);
        }
        Appearance.Row(layout, split, fill: true);
        Appearance.Row(layout, details);
        Controls.Add(layout);
        SetColumns();
    }

    internal async void LoadPreview()
    {
        if (versions is not null || loading) return;
        loading = true;
        var token = lifetime.Token;
        status.Text = "プレビューを読み込んでいます…";
        try
        {
            var result = await Task.Run(() =>
            {
                if (book.Plan is { } plan)
                    return (Versions: plan.ReadPreview(token), Errors: Array.Empty<string>());
                var errors = new List<string>();
                var paths = new[] { book.BasePath, book.ProductionPath, book.IncomingPath };
                var loaded = new Workbook?[paths.Length];
                for (var i = 0; i < paths.Length; i++)
                {
                    if (paths[i] is not string path) continue;
                    try { loaded[i] = XlsxReader.Read(path, token); }
                    catch (Exception error) when (FolderComparer.IsReadError(error)) { errors.Add($"{Titles[i]}を読み込めません。"); }
                }
                return (Versions: loaded, Errors: errors.ToArray());
            }, token);
            if (IsDisposed || token.IsCancellationRequested) return;
            versions = result.Versions;
            var names = versions.Where(workbook => workbook is not null).SelectMany(workbook => workbook!.Sheets)
                .Select(sheet => sheet.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            sheets.Items.AddRange(names);
            status.Text = loadMessage = result.Errors.Length > 0 ? string.Join(" ", result.Errors) :
                names.Length == 0 ? "表示できるシートがありません。" : "";
            if (pendingReveal is { } target && sheets.Items.Contains(target.Sheet)) Reveal(target);
            else if (names.Length > 0) sheets.SelectedIndex = 0;
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error))
        {
            if (!IsDisposed) status.Text = error.Message;
        }
        finally { loading = false; }
    }

    internal void SetOptions(PreviewOptions next)
    {
        var redraw = options.UseSavedZoom != next.UseSavedZoom || options.ShowGridlines != next.ShowGridlines;
        options = next;
        SetColumns();
        if (redraw) DrawSheet();
    }

    internal void Reveal(MergeCellAddress cell)
    {
        if (versions is null)
        {
            pendingReveal = cell;
            return;
        }
        pendingReveal = null;
        if (!sheets.Items.Contains(cell.Sheet)) return;
        selecting = true;
        sheets.SelectedItem = cell.Sheet;
        selecting = false;
        selection = cell;
        DrawSheet(jump: true);
    }

    private void SetColumns()
    {
        pairs.Visible = useBase && !options.ShowThreeWay;
        hiddenPane = !useBase ? 0 : options.ShowThreeWay ? -1 : pairs.SelectedIndex switch { 1 => 2, 2 => 1, _ => 0 };
        for (var i = 0; i < panes.Length; i++)
        {
            panels[i].Visible = i != hiddenPane;
            split.ColumnStyles[i].SizeType = i == hiddenPane ? SizeType.Absolute : SizeType.Percent;
            split.ColumnStyles[i].Width = i == hiddenPane ? 0 : hiddenPane < 0 ? 100 / 3f : 50;
        }
        if (sheets.SelectedItem is string name) LoadMarks(name);
    }

    private void DrawSheet(bool jump = false)
    {
        if (versions is null || sheets.SelectedItem is not string name) return;
        for (var i = 0; i < panes.Length; i++)
        {
            var pane = panes[i];
            var version = versions[i];
            var sheet = FindSheet(version, name);
            pane.SetGridlines(options.ShowGridlines);
            var merge = selection is { } cell
                ? sheet?.Merges.InRows(cell.Row, cell.Row).FirstOrDefault(range => range.ContainsColumn(cell.Column)) : null;
            pane.ShowSheet(sheet, selection?.Row ?? 1, selection?.Column ?? 1, merge, rowLayout: false,
                columnLayout: false, options.UseSavedZoom, version?.Date1904 ?? false, jump, selection is not null, 1);
        }
        LoadMarks(name);
    }

    private async void LoadMarks(string name)
    {
        if (versions is not { } books || markedScope == (name, hiddenPane)) return;
        markedScope = (name, hiddenPane);
        comparing?.Cancel();
        comparing = null;
        details.Clear();
        (string Sheet, int Left, int Right)[] targets = hiddenPane switch
        {
            -1 => [(name, 0, 1), (name, 0, 2)],
            2 => [(name, 0, 1)],
            1 => [(name, 0, 2)],
            _ => [(name, 1, 2)]
        };
        var missing = targets.Where(pair => !comparisons.ContainsKey(pair)).ToArray();
        if (missing.Length == 0)
        {
            ApplyMarks(name, CreateOverlay(targets));
            status.Text = loadMessage;
            return;
        }

        ApplyMarks(name, new PaneDifferences[3]);
        status.Text = "シートの差分を比較しています…";
        using var request = CancellationTokenSource.CreateLinkedTokenSource(lifetime.Token);
        comparing = request;
        try
        {
            var results = await Task.Run(() => missing.Select(pair =>
                CompareSheet(books[pair.Left], books[pair.Right], pair.Sheet, request.Token)).ToArray(), request.Token);
            if (IsDisposed || comparing != request) return;
            for (var i = 0; i < missing.Length; i++) comparisons[missing[i]] = results[i];
            ApplyMarks(name, CreateOverlay(targets));
            status.Text = loadMessage;
        }
        catch (OperationCanceledException) when (request.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error))
        {
            if (!IsDisposed && comparing == request)
            {
                markedScope = null;
                status.Text = $"シートの差分を表示できません。{error.Message}";
            }
        }
        finally { if (comparing == request) comparing = null; }
    }

    private PaneDifferences[] CreateOverlay((string Sheet, int Left, int Right)[] targets)
    {
        var overlay = new PaneDifferences[3];
        foreach (var pair in targets)
        {
            var differences = comparisons[pair];
            var earlier = overlay[pair.Left].Items;
            overlay[pair.Left] = new(earlier is null ? differences : earlier.Concat(differences).ToArray(), true);
            overlay[pair.Right] = new(differences, false);
        }
        return overlay;
    }

    private void ApplyMarks(string name, PaneDifferences[] overlay)
    {
        for (var i = 0; i < panes.Length; i++)
        {
            displayed[i] = overlay[i];
            panes[i].SetChanges(overlay[i].Items, overlay[i].Before);
            var related = overlay[i].Items;
            headings[i].Text = FindSheet(versions![i], name) is null ? $"{Titles[i]}　シートなし" :
                related?.Any(change => change.Item == "シート" && change.Kind is ChangeKind.Added or ChangeKind.Removed) == true
                    ? $"{Titles[i]}　シートの追加・削除あり" : Titles[i];
        }
    }

    private void ShowPicked(int index, FullSheetPane.CellPick pick)
    {
        var differences = displayed[index].Items;
        if (pick.DifferenceIndices.Count > 0 && differences is not null)
        {
            details.Text = $"{Titles[index]}　{pick.Address}{Environment.NewLine}" +
                string.Join(Environment.NewLine + Environment.NewLine, pick.DifferenceIndices
                    .Select(item => WorkbookView.DescribeDifference(differences[item])).Distinct());
            return;
        }
        var formula = pick.Formula is null ? "" : $"{Environment.NewLine}数式　={pick.Formula}";
        details.Text = $"{Titles[index]}　{pick.Address}　{pick.Value}{formula}";
    }

    private static Sheet? FindSheet(Workbook? book, string name) =>
        book?.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));

    private static IReadOnlyList<Difference> CompareSheet(Workbook? left, Workbook? right, string name, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (left is null || right is null) return [];
        var before = FindSheet(left, name);
        var after = FindSheet(right, name);
        return WorkbookComparer.Compare(new(before is null ? [] : [before], left.Date1904),
            new(after is null ? [] : [after], right.Date1904), token: token).SingleOrDefault()?.Differences ?? [];
    }

    private readonly record struct PaneDifferences(IReadOnlyList<Difference>? Items, bool Before);

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
