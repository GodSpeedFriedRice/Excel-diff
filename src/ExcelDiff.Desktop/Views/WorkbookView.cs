using System.Globalization;
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
    private readonly CheckBox only = new() { Text = "差分のある行だけ", AutoSize = true };
    private readonly CheckBox colors = new() { Text = "差分色", AutoSize = true, Checked = true };
    private readonly CheckBox savedZoom = new() { Text = "保存倍率", AutoSize = true };
    private readonly Label position = Appearance.Label("");
    private readonly Label warning = Appearance.Label("");
    private readonly DataGridView a = Appearance.Grid();
    private readonly DataGridView b = Appearance.Grid();
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage preview = new("プレビュー");
    private readonly TabPage xml = new("XML");
    private readonly TextBox xmlLeft = XmlBox();
    private readonly TextBox xmlRight = XmlBox();
    private readonly TextBox detail = new()
    {
        Dock = DockStyle.Fill, Height = 80, Multiline = true, ReadOnly = true,
        ScrollBars = ScrollBars.Vertical, BackColor = Color.White, AccessibleName = "変更前後の詳細"
    };
    private readonly Dictionary<(string, float, FontStyle), Font> fonts = [];
    private AlignedRow[] rows = [];
    private Dictionary<int, int> leftRows = [];
    private Dictionary<int, int> rightRows = [];
    private HashSet<(int, int)> changed = [];
    private int cursor = -1;
    private bool updating;
    private bool synchronizing;
    private bool comparing;

    internal string LeftPath { get; }
    internal string RightPath { get; }
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
        BackColor = Color.White;

        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        toolbar.Controls.AddRange([sheetPicker, only, colors, savedZoom,
            Appearance.Button("↑ 前へ", () => MoveDifference(-1)), position, Appearance.Button("次へ ↓", () => MoveDifference(1))]);
        foreach (var control in new Control[] { sheetPicker, only, colors, savedZoom }) control.Margin = new Padding(6, 8, 6, 4);
        Appearance.Row(layout, toolbar);

        var keyBar = Appearance.Toolbar();
        keyBar.Controls.AddRange([Appearance.Label("キー列（空欄はセル番地）"), key,
            Appearance.Button("行を対応付け", async () => await ApplyKey()), warning]);
        key.Margin = new Padding(4, 8, 4, 4);
        warning.ForeColor = Appearance.Muted;
        Appearance.Row(layout, keyBar);

        preview.Controls.Add(Split(Pane("左 ・ 変更前", leftPath, a), Pane("右 ・ 変更後", rightPath, b)));
        xml.Controls.Add(Split(xmlLeft, xmlRight));
        views.TabPages.AddRange([preview, xml]);
        Appearance.Row(layout, views, fill: true);
        var footer = Appearance.Stack();
        footer.AutoSize = true;
        var legend = Appearance.Toolbar();
        legend.Controls.AddRange([Legend("変更セル", Appearance.Changed), Legend("追加行", Appearance.Added),
            Legend("削除行", Appearance.Removed)]);
        Appearance.Row(footer, legend);
        Appearance.Row(footer, detail);
        Appearance.Row(layout, footer);
        Controls.Add(layout);

        SetupGrid(a, b, true);
        SetupGrid(b, a, false);
        sheetPicker.SelectedIndexChanged += (_, _) => ShowSheet();
        only.CheckedChanged += (_, _) => RenderRows();
        savedZoom.CheckedChanged += (_, _) => ShowSheet();
        colors.CheckedChanged += (_, _) => { a.Invalidate(); b.Invalidate(); };
        views.SelectedIndexChanged += (_, _) => UpdateXml();
        LoadPicker(0);
    }

    private static Label Legend(string text, Color color)
    {
        var label = Appearance.Label(text);
        label.BackColor = color;
        label.Padding = new Padding(6, 2, 6, 2);
        return label;
    }

    private static TextBox XmlBox() => new()
    {
        Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, WordWrap = false,
        ScrollBars = ScrollBars.Both, BackColor = Color.White
    };

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

    private static Control Pane(string name, string path, DataGridView grid)
    {
        var pane = Appearance.Stack();
        var label = Appearance.Label($"{name}     {path}");
        label.BackColor = Appearance.Panel;
        label.Dock = DockStyle.Fill;
        label.Margin = Padding.Empty;
        label.Padding = new Padding(8);
        Appearance.Row(pane, label);
        Appearance.Row(pane, grid, fill: true);
        return pane;
    }

    private void SetupGrid(DataGridView grid, DataGridView other, bool before)
    {
        grid.VirtualMode = true;
        grid.AllowUserToResizeColumns = false;
        grid.CellValueNeeded += (_, e) => FillCell(e, before);
        grid.CellFormatting += (_, e) => StyleCell(e, before);
        grid.CellClick += (_, e) =>
        {
            if (e.RowIndex < 0 || e.ColumnIndex <= 0 || Current is not SheetDiff sheet) return;
            var row = rows[e.RowIndex];
            for (var index = 0; index < sheet.Differences.Count; index++)
            {
                var change = sheet.Differences[index];
                if ((change.LeftRow is not null || change.RightRow is not null) &&
                    change.LeftRow == row.Left && change.RightRow == row.Right &&
                    (change.Column is null || change.Column == e.ColumnIndex))
                {
                    ShowDifference(index);
                    return;
                }
            }
            ShowCell(row, e.ColumnIndex);
        };
        grid.Scroll += (_, e) =>
        {
            if (synchronizing || updating || other.RowCount == 0) return;
            synchronizing = true;
            try
            {
                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll && grid.FirstDisplayedScrollingRowIndex >= 0)
                    other.FirstDisplayedScrollingRowIndex = grid.FirstDisplayedScrollingRowIndex;
                else if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll && grid.FirstDisplayedScrollingColumnIndex > 0)
                    other.FirstDisplayedScrollingColumnIndex = grid.FirstDisplayedScrollingColumnIndex;
            }
            finally { synchronizing = false; }
        };
        grid.Paint += (_, e) => PaintMerges(grid, e.Graphics, before);
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

    internal void SelectSheet(string name)
    {
        for (var index = 0; index < sheets.Count; index++)
        {
            if (!string.Equals(sheets[index].Name, name, StringComparison.OrdinalIgnoreCase)) continue;
            sheetPicker.SelectedIndex = index;
            return;
        }
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
        if (updating) return;
        cursor = -1;
        var current = Current;
        warning.Text = current?.KeyWarning ?? "";
        changed = current?.Differences.Where(d => d.Column.HasValue).SelectMany(d =>
        {
            var positions = new List<(int, int)>(2);
            if (d.LeftRow is int l) positions.Add((l, d.Column!.Value));
            if (d.RightRow is int r) positions.Add((-r, d.Column!.Value));
            return positions;
        }).ToHashSet() ?? [];

        var columns = new[] { MaxColumn(current?.Left), MaxColumn(current?.Right), 1 }.Max();
        updating = true;
        SetupColumns(a, current?.Left, columns);
        SetupColumns(b, current?.Right, columns);
        updating = false;
        RenderRows();
        UpdateXml();
        StatusChanged?.Invoke(this, EventArgs.Empty);
    }

    private static int MaxColumn(Sheet? sheet)
    {
        if (sheet is null) return 0;
        var count = sheet.Rows.Values.SelectMany(r => r.Keys).DefaultIfEmpty().Max();
        foreach (var setting in sheet.Settings.Keys)
        {
            if (setting.StartsWith("結合 ", StringComparison.Ordinal))
                count = Math.Max(count, Address.Parse(setting.Split(':')[^1].Split(' ')[^1]).Column);
        }
        return count;
    }

    private static int MergeEndRow(Sheet? sheet) => sheet?.Settings.Keys
        .Where(key => key.StartsWith("結合 ", StringComparison.Ordinal))
        .Select(key => Address.Parse(key.Split(':')[^1].Split(' ')[^1]).Row).DefaultIfEmpty().Max() ?? 0;

    private void SetupColumns(DataGridView grid, Sheet? sheet, int count)
    {
        grid.RowCount = 0;
        grid.Columns.Clear();
        var zoom = Zoom(sheet);
        grid.RowTemplate.Height = Math.Clamp((int)(Number(sheet, "既定レイアウト / defaultRowHeight", 20) * 1.33 * zoom), 12, 300);
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#", Width = 48, Frozen = true, SortMode = DataGridViewColumnSortMode.NotSortable });
        for (var i = 1; i <= count; i++)
        {
            var title = Address.Column(i);
            var width = Number(sheet, $"列 {title} / width", Number(sheet, "既定レイアウト / defaultColWidth", 13));
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = title,
                Width = Math.Clamp((int)((width * 7 + 5) * zoom), 24, 900), SortMode = DataGridViewColumnSortMode.NotSortable });
        }
    }

    private void RenderRows()
    {
        if (updating) return;
        var sheet = Current;
        var matches = sheet?.Rows ?? [];
        var positional = sheet is { UsesKey: false, Left: not null, Right: not null };
        if (only.Checked && sheet is not null)
        {
            var included = sheet.Differences.Where(d => d.LeftRow is not null || d.RightRow is not null)
                .Select(d => (d.LeftRow, d.RightRow)).ToHashSet();
            matches = matches.Where(row => included.Contains((row.Left, row.Right))).ToArray();
        }
        else if (positional)
        {
            // Keep blank rows in fixed-layout sheets without losing one-sided changes.
            var populated = matches.ToDictionary(row => (row.Left ?? row.Right)!.Value);
            var last = Math.Max(populated.Keys.DefaultIfEmpty().Max(),
                Math.Max(MergeEndRow(sheet!.Left), MergeEndRow(sheet.Right)));
            matches = Enumerable.Range(1, last).Select(i => populated.TryGetValue(i, out var row) ? row : new AlignedRow(i, i)).ToArray();
        }
        rows = matches.ToArray();
        leftRows = IndexRows(rows, before: true);
        rightRows = IndexRows(rows, before: false);
        updating = true;
        a.RowCount = b.RowCount = rows.Length;
        ApplyRowHeights(a, sheet?.Left, leftRows);
        ApplyRowHeights(b, sheet?.Right, rightRows);
        updating = false;
        if (sheet is { Differences.Count: > 0 }) ShowDifference(Math.Max(0, cursor));
        else
        {
            position.Text = "0 / 0";
            detail.Text = sheet is null ? "比較できるシートがありません。" : "比較対象の項目に差分はありません。";
        }
    }

    private static Dictionary<int, int> IndexRows(AlignedRow[] rows, bool before)
    {
        var index = new Dictionary<int, int>();
        for (var i = 0; i < rows.Length; i++)
            if ((before ? rows[i].Left : rows[i].Right) is int row) index[row] = i;
        return index;
    }

    private void ApplyRowHeights(DataGridView grid, Sheet? sheet, Dictionary<int, int> index)
    {
        if (sheet is null) return;
        foreach (var (name, value) in sheet.Settings)
        {
            const string suffix = " / ht";
            if (!name.StartsWith("行 ", StringComparison.Ordinal) || !name.EndsWith(suffix, StringComparison.Ordinal) ||
                !int.TryParse(name.AsSpan(2, name.Length - 2 - suffix.Length), out var row) ||
                !index.TryGetValue(row, out var gridIndex) || !double.TryParse(value, CultureInfo.InvariantCulture, out var height))
                continue;
            grid.Rows[gridIndex].Height = Math.Clamp((int)(height * 1.33 * Zoom(sheet)), 12, 400);
        }
    }

    private void FillCell(DataGridViewCellValueEventArgs e, bool before)
    {
        if (updating || e.RowIndex < 0 || e.RowIndex >= rows.Length) return;
        var row = before ? rows[e.RowIndex].Left : rows[e.RowIndex].Right;
        var sheet = before ? Current?.Left : Current?.Right;
        if (e.ColumnIndex == 0) e.Value = row?.ToString() ?? "—";
        else if (row is int r) e.Value = Display(sheet?.At(r, e.ColumnIndex), before ? left.Date1904 : right.Date1904);
        else if (e.ColumnIndex == 1) e.Value = "対応する行なし";
    }

    private void StyleCell(DataGridViewCellFormattingEventArgs e, bool before)
    {
        if (updating || e.RowIndex < 0 || e.RowIndex >= rows.Length || e.CellStyle is not DataGridViewCellStyle style) return;
        var pair = rows[e.RowIndex];
        var row = before ? pair.Left : pair.Right;
        var sheet = before ? Current?.Left : Current?.Right;
        if (e.ColumnIndex == 0 || row is null)
        {
            style.BackColor = Appearance.Panel;
            style.ForeColor = Appearance.Muted;
            return;
        }

        var cell = sheet?.At(row.Value, e.ColumnIndex);
        if (cell is not null)
        {
            style.Font = CellFont(cell, Zoom(sheet));
            style.ForeColor = ReadColor(cell.Format, "フォント/color/rgb") ?? Appearance.Ink;
            if (cell.Format.GetValueOrDefault("背景/patternFill/patternType") == "solid")
                style.BackColor = ReadColor(cell.Format, "背景/patternFill/fgColor/rgb") ?? Color.White;
            style.WrapMode = cell.Format.GetValueOrDefault("配置/wrapText") == "1" ? DataGridViewTriState.True : DataGridViewTriState.False;
            var alignment = cell.Format.GetValueOrDefault("配置/horizontal");
            if (alignment == "center") style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            else if (alignment == "right" || (alignment is null or "general" && cell.Kind == "n"))
                style.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        if (!colors.Checked) return;
        if (pair.Left is null) style.BackColor = Appearance.Added;
        else if (pair.Right is null) style.BackColor = Appearance.Removed;
        else if (changed.Contains((before ? row.Value : -row.Value, e.ColumnIndex))) style.BackColor = Appearance.Changed;
    }

    private void PaintMerges(DataGridView grid, Graphics graphics, bool before)
    {
        if (updating || only.Checked || grid.FirstDisplayedScrollingRowIndex < 0) return;
        var sheet = before ? Current?.Left : Current?.Right;
        if (sheet is null) return;
        var index = before ? leftRows : rightRows;
        foreach (var setting in sheet.Settings.Keys)
        {
            if (!setting.StartsWith("結合 ", StringComparison.Ordinal)) continue;
            var ends = setting[3..].Split(':', 2);
            var first = Address.Parse(ends[0]);
            var last = Address.Parse(ends[^1]);
            if (!index.TryGetValue(first.Row, out var start) || !index.TryGetValue(last.Row, out var end) ||
                end - start != last.Row - first.Row || last.Column >= grid.ColumnCount)
                continue;
            var firstVisible = Math.Max(start, grid.FirstDisplayedScrollingRowIndex);
            var lastVisible = Math.Min(end, grid.FirstDisplayedScrollingRowIndex + grid.DisplayedRowCount(true) - 1);
            var firstColumn = Math.Max(first.Column, grid.FirstDisplayedScrollingColumnIndex);
            var lastColumn = Math.Min(last.Column, grid.FirstDisplayedScrollingColumnIndex + grid.DisplayedColumnCount(true) - 1);
            if (firstVisible > lastVisible || firstColumn > lastColumn) continue;
            var top = grid.GetCellDisplayRectangle(firstColumn, firstVisible, true);
            var bottom = grid.GetCellDisplayRectangle(lastColumn, lastVisible, true);
            if (top.IsEmpty || bottom.IsEmpty) continue;
            var rectangle = Rectangle.Union(top, bottom);
            rectangle.Width--;
            rectangle.Height--;
            var cell = sheet.At(first.Row, first.Column);
            var background = cell is not null && cell.Format.GetValueOrDefault("背景/patternFill/patternType") == "solid"
                ? ReadColor(cell.Format, "背景/patternFill/fgColor/rgb") ?? Color.White : Color.White;
            if (colors.Checked && changed.Contains((before ? first.Row : -first.Row, first.Column))) background = Appearance.Changed;
            using var brush = new SolidBrush(background);
            graphics.FillRectangle(brush, rectangle);
            using var border = new Pen(Appearance.Line);
            graphics.DrawRectangle(border, rectangle);
            if (cell is not null)
                TextRenderer.DrawText(graphics, Display(cell, before ? left.Date1904 : right.Date1904),
                    CellFont(cell, Zoom(sheet)), Rectangle.Inflate(rectangle, -3, 0), Appearance.Ink,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }

    private Font CellFont(Cell cell, double zoom)
    {
        var format = cell.Format;
        var name = format.GetValueOrDefault("フォント/name/val") ?? Font.Name;
        var size = float.TryParse(format.GetValueOrDefault("フォント/sz/val"), CultureInfo.InvariantCulture, out var points) ? points : Font.Size;
        var style = FontStyle.Regular;
        foreach (var (flag, attribute) in new[] { (FontStyle.Bold, "b"), (FontStyle.Italic, "i"),
            (FontStyle.Underline, "u"), (FontStyle.Strikeout, "strike") })
        {
            var path = "フォント/" + attribute;
            if (format.GetValueOrDefault(path) == "1" || format.GetValueOrDefault(path + "/val") is "1" or "single" or "double")
                style |= flag;
        }
        var fontKey = (name, (float)Math.Clamp(size * zoom, 5, 90), style);
        if (!fonts.TryGetValue(fontKey, out var font))
        {
            try { font = new Font(fontKey.Item1, fontKey.Item2, fontKey.Item3); }
            catch (ArgumentException) { font = new Font(Font.FontFamily, fontKey.Item2, FontStyle.Regular); }
            fonts.Add(fontKey, font);
        }
        return font;
    }

    private static Color? ReadColor(IReadOnlyDictionary<string, string> values, string key)
    {
        if (!values.TryGetValue(key, out var hex) || hex.Length < 6 ||
            !int.TryParse(hex.AsSpan(hex.Length - 6), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return null;
        return Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
    }

    private double Zoom(Sheet? sheet) => savedZoom.Checked ? Math.Clamp(Number(sheet, "表示倍率", 100), 25, 300) / 100 : 1;
    private static double Number(Sheet? sheet, string key, double fallback) =>
        double.TryParse(sheet?.Settings.GetValueOrDefault(key), CultureInfo.InvariantCulture, out var value) && double.IsFinite(value)
            ? value : fallback;

    private static string Display(Cell? cell, bool date1904)
    {
        if (cell is null) return "";
        if (cell.Kind != "n" || !double.TryParse(cell.Value, CultureInfo.InvariantCulture, out var number)) return cell.Display;
        var format = cell.Format.GetValueOrDefault("表示形式");
        if (format is "組み込み 14" or "組み込み 15" or "組み込み 16" or "組み込み 17" or "組み込み 22")
        {
            try
            {
                if (!date1904 && number >= 60 && number < 61) return "1900/02/29";
                var serial = date1904 ? number + 1462 : number < 60 ? number + 1 : number;
                return DateTime.FromOADate(serial).ToString(format == "組み込み 22" ? "yyyy/MM/dd HH:mm" : "yyyy/MM/dd");
            }
            catch (ArgumentException) { return cell.Display; }
        }
        return format switch
        {
            "組み込み 1" => number.ToString("0"),
            "組み込み 2" => number.ToString("0.00"),
            "組み込み 3" => number.ToString("#,##0"),
            "組み込み 4" => number.ToString("#,##0.00"),
            "組み込み 9" => number.ToString("0%"),
            "組み込み 10" => number.ToString("0.00%"),
            _ => cell.Display
        };
    }

    private void UpdateXml()
    {
        if (views.SelectedTab != xml) return;
        xmlLeft.Text = Current?.Left?.Xml.ToString() ?? "";
        xmlRight.Text = Current?.Right?.Xml.ToString() ?? "";
    }

    internal void MoveDifference(int delta)
    {
        var count = Current?.Differences.Count ?? 0;
        if (count > 0) ShowDifference(Math.Clamp(cursor + delta, 0, count - 1));
    }

    private void ShowDifference(int index)
    {
        if (Current is not SheetDiff sheet || index < 0 || index >= sheet.Differences.Count) return;
        cursor = index;
        position.Text = $"{index + 1} / {sheet.Differences.Count}";
        var diff = sheet.Differences[index];
        var leftAddress = Address.Format(diff.LeftRow, diff.Column);
        var rightAddress = Address.Format(diff.RightRow, diff.Column);
        detail.Text = $"{diff.Item}   {leftAddress} → {rightAddress}{Environment.NewLine}" +
            $"変更前  {diff.Left ?? "—"}{Environment.NewLine}変更後  {diff.Right ?? "—"}";
        if (diff.LeftRow is null && diff.RightRow is null) return;
        var row = diff.LeftRow is int l ? leftRows.GetValueOrDefault(l, -1) : rightRows.GetValueOrDefault(diff.RightRow!.Value, -1);
        if (row < 0) return;
        SelectCell(a, row, diff.Column ?? 0);
        SelectCell(b, row, diff.Column ?? 0);
    }

    private static void SelectCell(DataGridView grid, int row, int column)
    {
        if (row >= grid.RowCount || column >= grid.ColumnCount) return;
        grid.CurrentCell = grid.Rows[row].Cells[column];
    }

    private void ShowCell(AlignedRow row, int column)
    {
        var before = row.Left is int l ? Current?.Left?.At(l, column) : null;
        var after = row.Right is int r ? Current?.Right?.At(r, column) : null;
        detail.Text = $"{Address.Format(row.Left, column)} → {Address.Format(row.Right, column)}{Environment.NewLine}" +
            $"変更前  {Display(before, left.Date1904)}{Environment.NewLine}変更後  {Display(after, right.Date1904)}";
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing)
        {
            foreach (var font in fonts.Values) font.Dispose();
            fonts.Clear();
        }
    }
}
