using ExcelDiff.Core.Comparison;

namespace ExcelDiff.Desktop.Views;

internal sealed class PackageDiffView : UserControl
{
    private readonly DataGridView parts = Appearance.Grid();
    private readonly GridSort sort;
    private readonly XmlDiffView xml = new();
    private readonly ThreeWayXmlView? tripleXml;
    private readonly Label summary = Appearance.Label("");
    private readonly ComboBox comparedSide = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 154 };
    private readonly TextBox details = new()
    {
        Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false,
        Height = 76, BackColor = Color.White, AccessibleName = "部品のサイズと SHA-256"
    };
    private readonly string? beforePath;
    private readonly string? afterPath;
    private readonly bool threeWay;
    private IReadOnlyList<PackagePartDifference> pairs = [];
    private IReadOnlyList<PackagePartMergeDifference> triples = [];
    private int[] partOrder = [];
    private CancellationTokenSource? pending;
    private string[] paths = [];
    private bool loaded;
    private bool updating;
    private bool showThreeWay;

    internal PackageDiffView(string? beforePath = null, string? afterPath = null)
    {
        this.beforePath = beforePath;
        this.afterPath = afterPath;
        threeWay = beforePath is null;
        if (threeWay) tripleXml = new ThreeWayXmlView { Visible = false };
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var header = Appearance.Toolbar();
        header.Controls.Add(summary);
        if (threeWay)
        {
            comparedSide.AccessibleName = "XMLの比較対象";
            comparedSide.Items.AddRange(["共通ベース ↔ 本番版", "共通ベース ↔ 取り込み元", "本番版 ↔ 取り込み元"]);
            comparedSide.SelectedIndex = 1;
            comparedSide.Margin = new Padding(6, 7, 6, 4);
            comparedSide.SelectedIndexChanged += (_, _) => ShowPart();
            header.Controls.Add(comparedSide);
        }
        Appearance.Row(layout, header);

        parts.VirtualMode = true;
        parts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        parts.AccessibleName = "ブックの部品差分";
        AddColumn("部品のパス", threeWay ? 47 : 65);
        if (threeWay)
        {
            AddColumn("本番版", 13);
            AddColumn("取り込み元", 13);
        }
        AddColumn("比較結果", 27);
        AddColumn("形式", 12);
        sort = new GridSort(parts, SortParts);
        parts.CellValueNeeded += (_, e) => e.Value = PartValue(PartIndex(e.RowIndex), e.ColumnIndex);
        parts.CellToolTipTextNeeded += (_, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 0) e.ToolTipText = PartValue(PartIndex(e.RowIndex), 0);
        };
        parts.CurrentCellChanged += (_, _) => ShowPart();

        var panes = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Margin = Padding.Empty };
        panes.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
        panes.RowStyles.Add(new RowStyle(SizeType.Percent, 65));
        panes.Controls.Add(parts, 0, 0);
        var detailPane = Appearance.Stack();
        Appearance.Row(detailPane, details);
        var code = new Panel { Dock = DockStyle.Fill };
        code.Controls.Add(xml);
        if (tripleXml is not null) code.Controls.Add(tripleXml);
        Appearance.Row(detailPane, code, fill: true);
        panes.Controls.Add(detailPane, 0, 1);
        Appearance.Row(layout, panes, fill: true);
        Controls.Add(layout);
        SetRows(0, threeWay ? "3冊を解析すると、ここに部品差分が表示されます。" : "");
        VisibleChanged += (_, _) =>
        {
            if (Visible) ShowPart();
            else
            {
                xml.Clear();
                tripleXml?.Clear();
            }
        };
    }

    internal void SetThreeWay(bool enabled)
    {
        if (!threeWay || showThreeWay == enabled) return;
        showThreeWay = enabled;
        comparedSide.Visible = !enabled;
        ShowPart();
    }

    private void AddColumn(string title, int width)
    {
        parts.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = title, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = width
        });
    }

    internal async void LoadComparison()
    {
        if (loaded || pending is not null || beforePath is null || afterPath is null) return;
        using var request = new CancellationTokenSource();
        pending = request;
        summary.Text = "ブック内の全部品を比較しています…";
        try
        {
            pairs = await Task.Run(() => PackageComparer.Compare(beforePath, afterPath, request.Token), request.Token);
            if (IsDisposed || request.IsCancellationRequested) return;
            loaded = true;
            SetRows(pairs.Count, pairs.Count == 0 ? "部品に差分はありません。" : $"変更のある部品 {pairs.Count} 件");
        }
        catch (OperationCanceledException) when (request.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error))
        {
            if (!IsDisposed) summary.Text = $"部品を比較できません。{error.Message}";
        }
        finally { pending = null; }
    }

    internal void SetMerge(IReadOnlyList<PackagePartMergeDifference> changes, string original, string production, string incoming)
    {
        triples = changes;
        paths = [original, production, incoming];
        var both = changes.Count(change => change.Status == PackagePartMergeStatus.BothChanged);
        SetRows(changes.Count, changes.Count == 0 ? "部品に差分はありません。" :
            $"変更のある部品 {changes.Count} 件　両側で別の変更 {both} 件");
    }

    private void SetRows(int count, string message)
    {
        updating = true;
        parts.RowCount = 0;
        partOrder = SortOrderOf(count);
        parts.RowCount = count;
        if (count > 0) parts.CurrentCell = parts.Rows[0].Cells[0];
        summary.Text = message;
        comparedSide.Enabled = count > 0;
        updating = false;
        if (count == 0)
        {
            xml.Clear();
            xml.Visible = false;
            details.Clear();
        }
        else ShowPart();
    }

    private int PartIndex(int row) => row >= 0 && row < partOrder.Length ? partOrder[row] : -1;

    private void SortParts()
    {
        var current = PartIndex(parts.CurrentCell?.RowIndex ?? -1);
        var column = parts.CurrentCell?.ColumnIndex ?? 0;
        partOrder = SortOrderOf(partOrder.Length);
        updating = true;
        if (current >= 0) parts.CurrentCell = parts[column, Array.IndexOf(partOrder, current)];
        updating = false;
        parts.Invalidate();
    }

    private int[] SortOrderOf(int count)
    {
        var order = Enumerable.Range(0, count);
        return (sort.Column switch
        {
            < 0 => order,
            1 when threeWay => sort.Order(order, index => triples[index].ProductionChange),
            2 when threeWay => sort.Order(order, index => triples[index].IncomingChange),
            3 when threeWay => sort.Order(order, index => (int)PackagePartMergeStatus.BothChanged - (int)triples[index].Status),
            1 => sort.Order(order, index => pairs[index].Kind),
            _ => sort.Order(order, index => PartValue(index, sort.Column), StringComparer.CurrentCultureIgnoreCase)
        }).ToArray();
    }

    private string? PartValue(int index, int column)
    {
        if (threeWay && index >= 0 && index < triples.Count)
        {
            var item = triples[index];
            return column switch
            {
                0 => item.Path,
                1 => ChangeLabel(item.ProductionChange),
                2 => ChangeLabel(item.IncomingChange),
                3 => item.Status switch
                {
                    PackagePartMergeStatus.ProductionOnly => "本番版のみ変更",
                    PackagePartMergeStatus.IncomingOnly => "取り込み元のみ変更",
                    PackagePartMergeStatus.SameChange => "両側で同じ変更",
                    _ => "両側で別の変更"
                },
                _ => Format(item.Base, item.Production, item.Incoming)
            };
        }
        if (!threeWay && index >= 0 && index < pairs.Count)
        {
            var item = pairs[index];
            return column switch { 0 => item.Path, 1 => ChangeLabel(item.Kind), _ => Format(item.Before, item.After) };
        }
        return null;
    }

    private void ShowPart()
    {
        if (updating || !Visible) return;
        xml.Clear();
        xml.Visible = false;
        if (tripleXml is not null)
        {
            tripleXml.Clear();
            tripleXml.Visible = false;
        }
        details.Clear();
        var index = PartIndex(parts.CurrentCell?.RowIndex ?? -1);
        if (index < 0) return;
        if (threeWay && index < triples.Count)
        {
            var item = triples[index];
            details.Lines = [Describe("共通ベース", item.Base), Describe("本番版", item.Production), Describe("取り込み元", item.Incoming)];
            if (showThreeWay && tripleXml is not null)
            {
                if (item.Base?.IsXml != false && item.Production?.IsXml != false && item.Incoming?.IsXml != false)
                {
                    tripleXml.Visible = true;
                    tripleXml.ShowPart(paths, item);
                }
                return;
            }
            var (left, right, leftPart, rightPart) = comparedSide.SelectedIndex switch
            {
                0 => (0, 1, item.Base, item.Production),
                2 => (1, 2, item.Production, item.Incoming),
                _ => (0, 2, item.Base, item.Incoming)
            };
            ShowXml(paths[left], leftPart, paths[right], rightPart);
        }
        else if (!threeWay && index < pairs.Count)
        {
            var item = pairs[index];
            details.Lines = [Describe("変更前", item.Before), Describe("変更後", item.After)];
            ShowXml(beforePath!, item.Before, afterPath!, item.After);
        }
    }

    private void ShowXml(string left, PackagePart? old, string right, PackagePart? updated)
    {
        if ((old is null && updated is null) || old?.IsXml == false || updated?.IsXml == false) return;
        xml.Visible = true;
        xml.ShowPart(left, old, right, updated);
    }

    private static string Describe(string label, PackagePart? part) => part is null ? $"{label}　なし" :
        $"{label}　{part.Length:N0} bytes　SHA-256 {part.Sha256}";

    private static string Format(params PackagePart?[] items) =>
        items.All(item => item?.IsXml != false) ? "XML" : items.Any(item => item?.IsXml == true) ? "混在" : "バイナリ";

    private static string ChangeLabel(PackagePartChange change) => change switch
    {
        PackagePartChange.Added => "追加",
        PackagePartChange.Removed => "削除",
        PackagePartChange.Modified => "変更",
        _ => "変更なし"
    };

    protected override void Dispose(bool disposing)
    {
        if (disposing) pending?.Cancel();
        base.Dispose(disposing);
    }
}
