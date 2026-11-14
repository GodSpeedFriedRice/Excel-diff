using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class WorkbookView : UserControl
{
    private readonly Workbook left;
    private readonly Workbook right;
    private IReadOnlyList<SheetDiff> sheets;
    private readonly ComboBox sheetPicker = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 168, AccessibleName = "シート" };
    private readonly TextBox key = new() { Width = 44, MaxLength = 3, CharacterCasing = CharacterCasing.Upper, AccessibleName = "キー列" };
    private PreviewOptions previewOptions = new();
    private readonly Label position = Appearance.Label("");
    private readonly Label warning = Appearance.Label("");
    private readonly DifferenceGrid differences = Appearance.ConfigureGrid(new DifferenceGrid());
    private readonly WheelZoom listZoom = new();
    private readonly WheelZoom fullZoom = new();
    private readonly Label zoomIndicator = Appearance.Label("100%");
    private Font? listZoomFont;
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage list = new("差分一覧");
    private readonly TabPage full = new("全体プレビュー");
    private readonly TabPage xml = new("XML差分");
    private readonly FullSheetPane leftFull = new();
    private readonly FullSheetPane rightFull = new();
    private readonly XmlDiffView xmlDiff = new();
    private readonly TextBox detail = new()
    {
        Dock = DockStyle.Fill, Height = 70, Multiline = true, ReadOnly = true,
        ScrollBars = ScrollBars.Vertical, BackColor = Color.White, AccessibleName = "変更前後の詳細"
    };
    private readonly TextBox fullDetails = new()
    {
        Dock = DockStyle.Fill, Height = 63, Multiline = true, ReadOnly = true,
        ScrollBars = ScrollBars.Vertical, BackColor = Color.White, AccessibleName = "全体プレビューで選択したセル"
    };
    private readonly SheetPreview leftPreview = new();
    private readonly SheetPreview rightPreview = new();
    private readonly TableLayoutPanel listLayout = Appearance.Stack();
    private readonly TableLayoutPanel previews;

    private SheetDiff? fullSheet;
    private Difference? fullDifference;
    private bool updating;
    private bool comparing;

    internal string LeftPath { get; }
    internal string RightPath { get; }
    internal string? SelectedSheet => Current?.Name;
    internal string Status => Current is SheetDiff sheet ? $"{sheet.Name}  ・  差分 {sheet.Differences.Count} 件" : "比較できるシートがありません";
    internal event EventHandler? StatusChanged;
    private SheetDiff? Current => sheetPicker.SelectedIndex is var index && index >= 0 && index < sheets.Count ? sheets[index] : null;

    internal WorkbookView(string leftPath, string rightPath, Workbook left, Workbook right, IReadOnlyList<SheetDiff> sheets)
    {
        LeftPath = leftPath;
        RightPath = rightPath;
        this.left = left;
        this.right = right;
        this.sheets = sheets;

        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        toolbar.Controls.AddRange([sheetPicker,
            Appearance.Button("↑ 前へ", () => MoveDifference(-1)), position, Appearance.Button("次へ ↓", () => MoveDifference(1)), zoomIndicator]);
        sheetPicker.Margin = new Padding(6, 8, 6, 4);
        Appearance.Row(layout, toolbar);

        var keyBar = Appearance.Toolbar();
        keyBar.Controls.AddRange([Appearance.Label("キー列（空欄はセル番地）"), key,
            Appearance.Button("行を対応付け", async () => await ApplyKey()), warning]);
        key.Margin = new Padding(4, 8, 4, 4);
        warning.ForeColor = Appearance.Muted;
        Appearance.Row(layout, keyBar);

        differences.VirtualMode = true;
        differences.ShowCellToolTips = false;
        differences.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        differences.AccessibleName = "シートの差分一覧";
        foreach (var (title, width) in new[] { ("場所", 20), ("変更項目", 22), ("変更前", 29), ("変更後", 29) })
        {
            differences.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = title, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = width, SortMode = DataGridViewColumnSortMode.NotSortable
            });
        }
        differences.RowHeightInfoNeeded += (_, e) => e.Height = (int)Math.Round(28 * listZoom.Scale);
        differences.ZoomRequested += ZoomList;
        differences.CellValueNeeded += FillDifference;
        differences.CellDoubleClick += (_, e) => { if (e.RowIndex >= 0) views.SelectedTab = full; };
        differences.CurrentCellChanged += (_, _) => { if (!updating) UpdateSelection(); };

        previews = Split(Pane("左 ・ 変更前", leftPath, leftPreview), Pane("右 ・ 変更後", rightPath, rightPreview));
        listLayout.Controls.Add(differences, 0, 0);
        listLayout.Controls.Add(detail, 0, 1);
        listLayout.Controls.Add(previews, 0, 2);
        listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 56));
        listLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
        listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 44));
        list.Controls.Add(listLayout);
        xml.Controls.Add(xmlDiff);
        var fullLayout = Appearance.Stack();
        var legend = Appearance.Toolbar();
        foreach (var (caption, background) in new[]
            { ("変更", Appearance.Changed), ("追加", Appearance.Added), ("削除", Appearance.Removed) })
        {
            var label = Appearance.Label(caption);
            label.BackColor = background;
            label.Padding = new Padding(5, 1, 5, 1);
            legend.Controls.Add(label);
        }
        Appearance.Row(fullLayout, legend);
        Appearance.Row(fullLayout, Split(Pane("左 ・ 変更前", leftPath, leftFull), Pane("右 ・ 変更後", rightPath, rightFull)), fill: true);
        Appearance.Row(fullLayout, fullDetails);
        full.Controls.Add(fullLayout);
        leftFull.Scrolled += rightFull.ScrollBy;
        rightFull.Scrolled += leftFull.ScrollBy;
        leftFull.ZoomRequested += ZoomFull;
        rightFull.ZoomRequested += ZoomFull;
        leftFull.Picked += pick => SelectPreviewCell(pick, before: true);
        rightFull.Picked += pick => SelectPreviewCell(pick, before: false);
        views.TabPages.AddRange([list, full, xml]);
        Appearance.Row(layout, views, fill: true);
        Controls.Add(layout);

        sheetPicker.SelectedIndexChanged += (_, _) => { if (!updating) ShowSheet(); };
        views.SelectedIndexChanged += (_, _) =>
        {
            UpdateZoomIndicator();
            UpdateSelection();
            if (views.SelectedTab == full) leftFull.FocusSurface();
            else if (views.SelectedTab == list) differences.Focus();
        };
        LoadPicker(0);
    }

    internal void SetPreviewOptions(PreviewOptions options)
    {
        var previous = previewOptions;
        if (options == previous) return;
        previewOptions = options;
        if (options.ShowGridlines != previous.ShowGridlines)
        {
            leftPreview.SetGridlines(options.ShowGridlines);
            rightPreview.SetGridlines(options.ShowGridlines);
            leftFull.SetGridlines(options.ShowGridlines);
            rightFull.SetGridlines(options.ShowGridlines);
        }
        if (options.ShowNearby != previous.ShowNearby)
        {
            previews.Visible = options.ShowNearby;
            listLayout.RowStyles[2].SizeType = options.ShowNearby ? SizeType.Percent : SizeType.Absolute;
            listLayout.RowStyles[2].Height = options.ShowNearby ? 44 : 0;
        }
        if (options.UseSavedZoom != previous.UseSavedZoom || views.SelectedTab == list && options.ShowNearby != previous.ShowNearby)
            UpdateSelection(jumpToDifference: false);
    }

    private void ZoomList(int delta)
    {
        if (!listZoom.Apply(delta)) return;
        var font = new Font(Font.FontFamily, (float)(Font.SizeInPoints * listZoom.Scale), Font.Style, GraphicsUnit.Point);
        var previous = listZoomFont;
        listZoomFont = font;
        differences.SuspendLayout();
        try
        {
            differences.Font = font;
            differences.ColumnHeadersDefaultCellStyle.Font = font;
            detail.Font = font;
            differences.RowTemplate.Height = (int)Math.Round(28 * listZoom.Scale);
            differences.ColumnHeadersHeight = (int)Math.Round(30 * listZoom.Scale);
            listLayout.RowStyles[1].Height = (float)(76 * listZoom.Scale);
            differences.RefreshRowHeight();
        }
        finally { differences.ResumeLayout(); }
        previous?.Dispose();
        UpdateZoomIndicator();
        if (views.SelectedTab == list) UpdateSelection();
    }

    private void ZoomFull(int delta)
    {
        if (!fullZoom.Apply(delta)) return;
        leftFull.SetZoom(fullZoom.Scale);
        rightFull.SetZoom(fullZoom.Scale);
        UpdateZoomIndicator();
    }

    private void UpdateZoomIndicator()
    {
        zoomIndicator.Visible = views.SelectedTab != xml;
        zoomIndicator.Text = $"{(views.SelectedTab == full ? fullZoom.Percent : listZoom.Percent)}%";
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing) listZoomFont?.Dispose();
    }

    internal void SelectSheet(string name)
    {
        for (var index = 0; index < sheets.Count; index++)
        {
            if (!string.Equals(sheets[index].Name, name, StringComparison.OrdinalIgnoreCase)) continue;
            sheetPicker.SelectedIndex = index;
            return;
        }
    }


    private static TableLayoutPanel Split(Control left, Control right)
    {
        var split = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = Appearance.Line };
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        left.Margin = Padding.Empty;
        right.Margin = new Padding(1, 0, 0, 0);
        split.Controls.Add(left, 0, 0);
        split.Controls.Add(right, 1, 0);
        return split;
    }

    private static Control Pane(string title, string path, Control preview)
    {
        var pane = Appearance.Stack();
        var label = Appearance.Label($"{title}　{Path.GetFileName(path)}");
        Appearance.Row(pane, label);
        if (preview is SheetPreview)
        {
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            scroll.Controls.Add(preview);
            Appearance.Row(pane, scroll, fill: true);
        }
        else
            Appearance.Row(pane, preview, fill: true);
        return pane;
    }

    private async Task ApplyKey()
    {
        if (comparing) return;
        int? column = null;
        if (key.Text.Length != 0)
        {
            if (!Address.TryColumn(key.Text, out var number))
            {
                warning.Text = "キー列は A～XFD の列名で指定してください。";
                return;
            }
            column = number;
        }
        comparing = true;
        warning.Text = "行を対応付けています…";
        try
        {
            var result = await Task.Run(() => WorkbookComparer.Compare(left, right, column));
            if (IsDisposed) return;
            sheets = result;
            LoadPicker(Math.Max(0, sheetPicker.SelectedIndex));
        }
        finally { comparing = false; }
    }

    private void LoadPicker(int index)
    {
        updating = true;
        sheetPicker.Items.Clear();
        sheetPicker.Items.AddRange(sheets.Select(sheet => $"{sheet.Name}  {sheet.Differences.Count}").ToArray());
        sheetPicker.SelectedIndex = sheets.Count == 0 ? -1 : index;
        updating = false;
        ShowSheet();
    }

    private void ShowSheet()
    {
        warning.Text = Current?.KeyWarning ?? "";
        updating = true;
        differences.RowCount = Current?.Differences.Count ?? 0;
        if (differences.RowCount > 0)
        {
            differences.CurrentCell = differences[0, 0];
            if (differences.FirstDisplayedScrollingRowIndex > 0 && differences.DisplayedRowCount(false) > 0)
                differences.FirstDisplayedScrollingRowIndex = 0;
        }
        differences.Invalidate();
        updating = false;
        UpdateSelection();
        StatusChanged?.Invoke(this, EventArgs.Empty);
    }

    private static string DescribeLocation(Difference diff) => diff.LeftAddress == diff.RightAddress
        ? diff.LeftAddress : $"{diff.LeftAddress} → {diff.RightAddress}";

    private void FillDifference(object? sender, DataGridViewCellValueEventArgs e)
    {
        if (Current is not SheetDiff sheet || e.RowIndex < 0 || e.RowIndex >= sheet.Differences.Count) return;
        var diff = sheet.Differences[e.RowIndex];
        e.Value = e.ColumnIndex switch
        {
            0 => DescribeLocation(diff),
            1 => diff.Parts is { Count: > 0 } parts ? $"{diff.Item} ({parts.Count} 項目)" : diff.Item,
            2 => diff.Left ?? "—",
            _ => diff.Right ?? "—"
        };
    }

    private static string DescribeDifference(Difference diff)
    {
        var heading = $"{diff.Item}　{DescribeLocation(diff)}";
        return diff.Parts is { Count: > 0 } parts
            ? heading + Environment.NewLine + string.Join(Environment.NewLine,
                parts.Select(part => $"{part.Item}　{part.Left ?? "—"} → {part.Right ?? "—"}"))
            : $"{heading}{Environment.NewLine}変更前　{diff.Left ?? "—"}{Environment.NewLine}変更後　{diff.Right ?? "—"}";
    }

    private void SelectPreviewCell(FullSheetPane.CellPick pick, bool before)
    {
        if (Current is not SheetDiff sheet) return;
        if (pick.DifferenceIndices.Count > 0)
        {
            updating = true;
            differences.CurrentCell = differences[0, pick.DifferenceIndices[0]];
            updating = false;
            UpdateSelection(jumpToDifference: false);
            fullDetails.Text = string.Join(Environment.NewLine + Environment.NewLine,
                pick.DifferenceIndices.Select(index => DescribeDifference(sheet.Differences[index])));
            return;
        }
        (before ? rightFull : leftFull).ClearSelection();
        var side = before ? "左" : "右";
        var value = pick.Value.Length > 0 ? $"　{pick.Value}" : "";
        var formula = pick.Formula is null ? "" : $"{Environment.NewLine}数式　={pick.Formula}";
        fullDetails.Text = $"{side}　{pick.Address}{value}{formula}";
    }

    private void UpdateSelection(bool jumpToDifference = true)
    {
        var sheet = Current;
        if (views.SelectedTab == xml)
        {
            xmlDiff.ShowSheets(sheet?.Left, sheet?.Right);
            return;
        }

        var index = differences.CurrentCell?.RowIndex ?? -1;
        var count = sheet?.Differences.Count ?? 0;
        var diff = index >= 0 && index < count ? sheet!.Differences[index] : null;
        position.Text = $"{(diff is null ? 0 : index + 1)} / {count}";
        detail.Text = diff is not null ? DescribeDifference(diff) : sheet is null
            ? "比較できるシートがありません。" : "比較対象の項目に差分はありません。";

        var column = diff?.Column ?? 1;
        var leftRow = diff?.LeftRow ?? diff?.RightRow ?? 1;
        var rightRow = diff?.RightRow ?? diff?.LeftRow ?? 1;
        if (views.SelectedTab == full)
        {
            var jump = jumpToDifference && (!ReferenceEquals(fullSheet, sheet) || !ReferenceEquals(fullDifference, diff));
            fullDetails.Text = detail.Text;
            leftFull.SetChanges(sheet?.Differences, before: true);
            rightFull.SetChanges(sheet?.Differences, before: false);
            fullSheet = sheet;
            fullDifference = diff;
            var rowLayout = diff?.Kind == ChangeKind.RowLayout || diff?.Column is null && diff?.LeftRow is not null;
            var columnLayout = diff?.Kind == ChangeKind.ColumnLayout;
            var leftMark = diff?.LeftRow is not null || columnLayout;
            var rightMark = diff?.RightRow is not null || columnLayout;
            leftFull.ShowSheet(sheet?.Left, leftRow, column, diff?.LeftMerge, rowLayout, columnLayout,
                previewOptions.UseSavedZoom, left.Date1904, jump, leftMark, diff?.LayoutCount ?? 1);
            rightFull.ShowSheet(sheet?.Right, rightRow, column, diff?.RightMerge, rowLayout, columnLayout,
                previewOptions.UseSavedZoom, right.Date1904, jump, rightMark, diff?.LayoutCount ?? 1);
            return;
        }
        if (!previewOptions.ShowNearby) return;
        leftPreview.ShowArea(sheet?.Left, leftRow, column, diff?.LeftMerge, diff?.LeftRow is not null,
            previewOptions.UseSavedZoom, left.Date1904, listZoom.Scale);
        rightPreview.ShowArea(sheet?.Right, rightRow, column, diff?.RightMerge, diff?.RightRow is not null,
            previewOptions.UseSavedZoom, right.Date1904, listZoom.Scale);
    }


    internal void MoveDifference(int delta)
    {
        if (differences.RowCount == 0) return;
        if (views.SelectedTab == xml) views.SelectedTab = list;
        var index = Math.Clamp((differences.CurrentCell?.RowIndex ?? -1) + delta, 0, differences.RowCount - 1);
        differences.CurrentCell = differences[0, index];
    }
}
