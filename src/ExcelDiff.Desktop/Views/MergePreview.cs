using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergePreview : UserControl
{
    private readonly FolderMergeBook book;
    private readonly CancellationTokenSource lifetime = new();
    private readonly ComboBox sheets = new() { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "プレビューするシート" };
    private readonly ComboBox pairs = new() { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList, AccessibleName = "プレビューする版" };
    private readonly Label status = Appearance.Label("");
    private readonly TableLayoutPanel split = new() { ColumnCount = 3, RowCount = 1, Dock = DockStyle.Fill };
    private readonly FullSheetPane[] panes = [new(), new(), new()];
    private readonly Control[] panels = new Control[3];
    private readonly Label[] headings = new Label[3];
    private readonly WheelZoom zoom = new();
    private readonly Dictionary<string, MergeCellChange[]> changes;
    private PreviewOptions options = new();
    private Workbook?[]? versions;
    private MergeCellAddress? pendingReveal;
    private MergeCellAddress? selection;
    private string? markedSheet;
    private bool loading;
    private bool selecting;

    internal MergePreview(FolderMergeBook book)
    {
        this.book = book;
        changes = (book.Plan?.Changes ?? []).GroupBy(change => change.Cell.Sheet)
            .ToDictionary(group => group.Key, group => group.ToArray());
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
        var titles = new[] { "共通ベース", "取り込み1", "取り込み2" };
        for (var i = 0; i < panes.Length; i++)
        {
            var pane = panes[i];
            var panel = Appearance.Stack();
            var heading = Appearance.Label(titles[i]);
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
        }
        Appearance.Row(layout, split, fill: true);
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
                var names = new[] { "共通ベース", "取り込み1", "取り込み2" };
                var paths = new[] { book.BasePath, book.ProductionPath, book.IncomingPath };
                var loaded = new Workbook?[paths.Length];
                for (var i = 0; i < paths.Length; i++)
                {
                    if (paths[i] is not string path) continue;
                    try { loaded[i] = XlsxReader.Read(path, token); }
                    catch (Exception error) when (FolderComparer.IsReadError(error)) { errors.Add($"{names[i]}を読み込めません。"); }
                }
                return (Versions: loaded, Errors: errors.ToArray());
            }, token);
            if (IsDisposed || token.IsCancellationRequested) return;
            versions = result.Versions;
            var names = versions.Where(workbook => workbook is not null).SelectMany(workbook => workbook!.Sheets)
                .Select(sheet => sheet.Name).Distinct(StringComparer.Ordinal).ToArray();
            sheets.Items.AddRange(names);
            status.Text = result.Errors.Length > 0 ? string.Join(" ", result.Errors) :
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
        pairs.Visible = !options.ShowThreeWay;
        var hidden = options.ShowThreeWay ? -1 : pairs.SelectedIndex;
        // Pair choices omit base, incoming, then production.
        hidden = hidden switch { 1 => 2, 2 => 1, _ => hidden };
        for (var i = 0; i < panes.Length; i++)
        {
            panels[i].Visible = i != hidden;
            split.ColumnStyles[i].SizeType = i == hidden ? SizeType.Absolute : SizeType.Percent;
            split.ColumnStyles[i].Width = i == hidden ? 0 : hidden < 0 ? 100 / 3f : 50;
        }
    }

    private void DrawSheet(bool jump = false)
    {
        if (versions is null || sheets.SelectedItem is not string name) return;
        var titles = new[] { "共通ベース", "取り込み1", "取り込み2" };
        var sheetChanges = changes.GetValueOrDefault(name) ?? [];
        for (var i = 0; i < panes.Length; i++)
        {
            var pane = panes[i];
            var version = versions[i];
            var sheet = version?.Sheets.FirstOrDefault(item => item.Name == name);
            headings[i].Text = sheet is null ? $"{titles[i]}　シートなし" : titles[i];
            pane.SetGridlines(options.ShowGridlines);
            if (markedSheet != name)
            {
                var marks = sheetChanges.Where(change => i switch
                {
                    1 => change.Disposition is not MergeDisposition.TakeIncoming,
                    2 => change.Disposition is not MergeDisposition.KeepProduction,
                    _ => true
                }).Select(change => new Difference(ChangeKind.Value, "値・数式", null, null,
                    change.Cell.Row, change.Cell.Row, change.Cell.Column)).ToArray();
                pane.SetChanges(marks, before: i == 0);
            }
            var merge = selection is { } cell
                ? sheet?.Merges.InRows(cell.Row, cell.Row).FirstOrDefault(range => range.ContainsColumn(cell.Column)) : null;
            pane.ShowSheet(sheet, selection?.Row ?? 1, selection?.Column ?? 1, merge, rowLayout: false,
                columnLayout: false, options.UseSavedZoom, version?.Date1904 ?? false, jump, selection is not null, 1);
        }
        markedSheet = name;
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
