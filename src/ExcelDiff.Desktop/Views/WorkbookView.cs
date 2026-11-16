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
    private readonly GridSort differenceSort;
    private readonly WheelZoom listZoom = new();
    private readonly WheelZoom fullZoom = new();
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage list = new("差分一覧");
    private readonly TabPage full = new("全体プレビュー");
    private readonly TabPage xml = new("XML差分");
    private readonly TabPage package = new("ブック全体の差分");
    private readonly FullSheetPane leftFull = new();
    private readonly FullSheetPane rightFull = new();
    private readonly XmlDiffView xmlDiff = new();
    private readonly PackageDiffView packageDiff;
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
    private readonly FullSheetPane leftPreview = new();
    private readonly FullSheetPane rightPreview = new();
    private readonly TableLayoutPanel listLayout = Appearance.Stack();
    private readonly TableLayoutPanel previews;

    private int differenceIndex = -1;
    private int[] differenceOrder = [];
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
        packageDiff = new PackageDiffView(leftPath, rightPath);

        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        toolbar.Controls.AddRange([sheetPicker,
            Appearance.Button("↑ 前へ", () => MoveDifference(-1)), position, Appearance.Button("次へ ↓", () => MoveDifference(1))]);
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
                FillWeight = width
            });
        }
        differenceSort = new GridSort(differences, SortDifferences);
        differences.ZoomRequested += listZoom.Apply;
        differences.CellValueNeeded += FillDifference;
        differences.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex < 0) return;
            views.SelectedTab = full;
            SelectDifference(e.RowIndex, reveal: true);
        };
        differences.CurrentCellChanged += (_, _) =>
        {
            if (updating || views.SelectedTab != list) return;
            differenceIndex = differences.CurrentCell?.RowIndex ?? -1;
            UpdateSelection(reveal: true);
        };

        previews = Split(Pane("左 ・ 変更前", leftPath, leftPreview), Pane("右 ・ 変更後", rightPath, rightPreview));
        listLayout.Controls.Add(differences, 0, 0);
        listLayout.Controls.Add(detail, 0, 1);
        listLayout.Controls.Add(previews, 0, 2);
        listLayout.Controls.Add(new ZoomBar(listZoom), 0, 3);
        listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 56));
        listLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
        listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 44));
        listLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        list.Controls.Add(listLayout);
        xml.Controls.Add(xmlDiff);
        package.Controls.Add(packageDiff);
        var fullLayout = Appearance.Stack();
        Appearance.Row(fullLayout, Split(Pane("左 ・ 変更前", leftPath, leftFull), Pane("右 ・ 変更後", rightPath, rightFull)), fill: true);
        Appearance.Row(fullLayout, fullDetails);
        Appearance.Row(fullLayout, new ZoomBar(fullZoom));
        full.Controls.Add(fullLayout);
        leftFull.Scrolled += rightFull.ScrollBy;
        rightFull.Scrolled += leftFull.ScrollBy;
        leftPreview.Scrolled += rightPreview.ScrollBy;
        rightPreview.Scrolled += leftPreview.ScrollBy;
        leftFull.ZoomRequested += fullZoom.Apply;
        rightFull.ZoomRequested += fullZoom.Apply;
        leftPreview.ZoomRequested += listZoom.Apply;
        rightPreview.ZoomRequested += listZoom.Apply;
        fullZoom.Changed += () =>
        {
            leftFull.SetZoom(fullZoom.Scale);
            rightFull.SetZoom(fullZoom.Scale);
        };
        listZoom.Changed += ApplyListZoom;
        leftFull.Picked += pick => SelectPreviewCell(pick, before: true);
        rightFull.Picked += pick => SelectPreviewCell(pick, before: false);
        leftPreview.Picked += pick => SelectPreviewCell(pick, before: true);
        rightPreview.Picked += pick => SelectPreviewCell(pick, before: false);
        views.TabPages.AddRange([full, list, xml, package]);
        views.SelectedTab = full;
        Appearance.Row(layout, views, fill: true);
        Controls.Add(layout);

        sheetPicker.SelectedIndexChanged += (_, _) => { if (!updating) ShowSheet(); };
        views.SelectedIndexChanged += (_, _) =>
        {
            if (views.SelectedTab == list && differences.RowCount > 0)
                SelectDifference(Math.Max(0, differenceIndex), reveal: true);
            else
                UpdateSelection();
            if (views.SelectedTab == full) leftFull.FocusSurface();
            else if (views.SelectedTab == list) differences.Focus();
        };
        LoadPicker(0);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (views.SelectedTab == full)
        {
            if (keyData is (Keys.Control | Keys.Up) or (Keys.Control | Keys.Down))
            {
                MoveDifference(keyData == (Keys.Control | Keys.Up) ? -1 : 1);
                return true;
            }
            if ((keyData is Keys.Up or Keys.Down) && (leftFull.ContainsFocus || rightFull.ContainsFocus))
                return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
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
            UpdateSelection(reveal: options.ShowNearby && !previous.ShowNearby);
    }

    private void ApplyListZoom()
    {
        differences.SetZoom(listZoom.Scale);
        detail.Font = differences.Font;
        listLayout.RowStyles[1].Height = (float)Math.Clamp(76 * listZoom.Scale, 28, 120);
        leftPreview.SetZoom(listZoom.Scale);
        rightPreview.SetZoom(listZoom.Scale);
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
        differenceOrder = OrderedDifferences();
        differences.RowCount = differenceOrder.Length;
        differenceIndex = differences.RowCount > 0 && views.SelectedTab == list ? 0 : -1;
        if (differences.RowCount > 0)
        {
            differences.CurrentCell = differenceIndex >= 0 ? differences[0, differenceIndex] : null;
            if (differences.FirstDisplayedScrollingRowIndex > 0 && differences.DisplayedRowCount(false) > 0)
                differences.FirstDisplayedScrollingRowIndex = 0;
        }
        differences.Invalidate();
        updating = false;
        UpdateSelection(reveal: views.SelectedTab == list);
        StatusChanged?.Invoke(this, EventArgs.Empty);
    }

    private static string DescribeLocation(Difference diff) => diff.LeftAddress == diff.RightAddress
        ? diff.LeftAddress : $"{diff.LeftAddress} → {diff.RightAddress}";

    private int[] OrderedDifferences()
    {
        if (Current is not { } sheet) return [];
        var items = sheet.Differences;
        var indices = Enumerable.Range(0, items.Count);
        var text = StringComparer.CurrentCultureIgnoreCase;
        return (differenceSort.Column switch
        {
            0 => differenceSort.Order(indices, index =>
            {
                var diff = items[index];
                return (diff.LeftRow ?? diff.RightRow ?? 0, diff.Column ?? 0);
            }),
            1 => differenceSort.Order(indices, index => items[index].Item, text),
            2 => differenceSort.Order(indices, index => items[index].Left, text),
            3 => differenceSort.Order(indices, index => items[index].Right, text),
            _ => indices
        }).ToArray();
    }

    private void SortDifferences()
    {
        var selected = DifferenceAt(differenceIndex);
        var column = differences.CurrentCell?.ColumnIndex ?? 0;
        differenceOrder = OrderedDifferences();
        differenceIndex = selected < 0 ? -1 : Array.IndexOf(differenceOrder, selected);
        updating = true;
        if (differenceIndex >= 0) differences.CurrentCell = differences[column, differenceIndex];
        updating = false;
        differences.Invalidate();
        UpdateSelection();
    }

    private int DifferenceAt(int row) => row >= 0 && row < differenceOrder.Length ? differenceOrder[row] : -1;

    private void FillDifference(object? sender, DataGridViewCellValueEventArgs e)
    {
        var index = DifferenceAt(e.RowIndex);
        if (Current is not SheetDiff sheet || index < 0) return;
        var diff = sheet.Differences[index];
        e.Value = e.ColumnIndex switch
        {
            0 => DescribeLocation(diff),
            1 => diff.Parts is { Count: > 0 } parts ? $"{diff.Item} ({parts.Count} 項目)" : diff.Item,
            2 => diff.Left ?? "—",
            _ => diff.Right ?? "—"
        };
    }

    internal static string DescribeDifference(Difference diff)
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
        var nearby = views.SelectedTab == list;
        var details = nearby ? detail : fullDetails;
        if (pick.DifferenceIndices.Count > 0)
        {
            var displayIndex = Array.IndexOf(differenceOrder, pick.DifferenceIndices[0]);
            if (displayIndex >= 0) SelectDifference(displayIndex);
            details.Text = string.Join(Environment.NewLine + Environment.NewLine,
                pick.DifferenceIndices.Select(index => DescribeDifference(sheet.Differences[index])));
            return;
        }
        var opposite = nearby ? (before ? rightPreview : leftPreview) : (before ? rightFull : leftFull);
        opposite.ClearSelection();
        var side = before ? "左" : "右";
        var value = pick.Value.Length > 0 ? $"　{pick.Value}" : "";
        var formula = pick.Formula is null ? "" : $"{Environment.NewLine}数式　={pick.Formula}";
        details.Text = $"{side}　{pick.Address}{value}{formula}";
    }

    private void UpdateSelection(bool reveal = false)
    {
        if (views.SelectedTab == package)
        {
            packageDiff.LoadComparison();
            return;
        }
        var sheet = Current;
        if (views.SelectedTab == xml)
        {
            xmlDiff.ShowSheets(sheet?.Left, sheet?.Right);
            return;
        }

        var index = differenceIndex;
        var count = sheet?.Differences.Count ?? 0;
        var originalIndex = DifferenceAt(index);
        var diff = sheet is not null && originalIndex >= 0 ? sheet.Differences[originalIndex] : null;
        position.Text = $"{(diff is null ? 0 : index + 1)} / {count}";
        detail.Text = diff is not null ? DescribeDifference(diff) : sheet is null
            ? "比較できるシートがありません。" : count == 0 ? "比較対象の項目に差分はありません。" :
                "セルを選択すると詳細を表示します。次の差分へは Ctrl+↓ で移動できます。";

        var column = diff?.Column ?? 1;
        var leftRow = diff?.LeftRow ?? diff?.RightRow ?? 1;
        var rightRow = diff?.RightRow ?? diff?.LeftRow ?? 1;
        var wholeSheet = views.SelectedTab == full;
        if (wholeSheet) fullDetails.Text = detail.Text;
        else if (!previewOptions.ShowNearby) return;
        var beforePane = wholeSheet ? leftFull : leftPreview;
        var afterPane = wholeSheet ? rightFull : rightPreview;
        beforePane.SetChanges(sheet?.Differences, before: true);
        afterPane.SetChanges(sheet?.Differences, before: false);
        var rowLayout = diff?.Kind == ChangeKind.RowLayout || diff?.Column is null && diff?.LeftRow is not null;
        var columnLayout = diff?.Kind == ChangeKind.ColumnLayout;
        var leftMark = diff?.LeftRow is not null || columnLayout;
        var rightMark = diff?.RightRow is not null || columnLayout;
        beforePane.ShowSheet(sheet?.Left, leftRow, column, diff?.LeftMerge, rowLayout, columnLayout,
            previewOptions.UseSavedZoom, left.Date1904, reveal && diff is not null, leftMark, diff?.LayoutCount ?? 1);
        afterPane.ShowSheet(sheet?.Right, rightRow, column, diff?.RightMerge, rowLayout, columnLayout,
            previewOptions.UseSavedZoom, right.Date1904, reveal && diff is not null, rightMark, diff?.LayoutCount ?? 1);
    }


    internal void MoveDifference(int delta)
    {
        if (differences.RowCount == 0) return;
        var index = Math.Clamp(differenceIndex + delta, 0, differences.RowCount - 1);
        if (views.SelectedTab == xml) views.SelectedTab = list;
        SelectDifference(index, reveal: true);
    }

    private void SelectDifference(int index, bool reveal = false)
    {
        differenceIndex = index;
        updating = true;
        differences.CurrentCell = differences[0, index];
        updating = false;
        UpdateSelection(reveal);
    }
}
