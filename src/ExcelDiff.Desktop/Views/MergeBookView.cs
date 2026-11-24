using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeBookView : UserControl
{
    private readonly DifferenceGrid grid = new();
    private readonly WheelZoom zoom = new();
    private readonly GridSort sort;
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage sheetTab = new("シート差分");
    private readonly TabPage cellTab = new("差分・コンフリクト");
    private readonly TabPage previewTab = new("全体プレビュー");
    private readonly Label status = Appearance.Label("");
    private readonly TextBox details = new()
    {
        Dock = DockStyle.Top, Height = 56, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
        AccessibleName = "選択した差分の詳細"
    };
    private readonly PackageDiffView? package;
    private readonly MergePreview preview;
    private readonly MergeComparisonView comparison;
    private readonly Button useProduction;
    private readonly Button useIncoming;
    private readonly Button reveal;
    private readonly Button nextUnresolved;
    private MergeCellChange[] rows;
    private readonly IReadOnlyDictionary<MergeCellAddress, MergeChoice> choices;
    private readonly Dictionary<MergeCellAddress, MergeCellChange> changesByCell;
    private readonly (MergeRange Range, MergeCellChange Change)[] mergeChanges;
    private readonly MergeCellChange[] conflicts;
    private MergeCellChange? previewChange;
    private MergeCellChange? comparisonChange;
    private SheetDiff? comparisonSheet;
    private Difference? comparisonDifference;
    private readonly string notes;
    private bool locked;

    internal MergeBookView(MergeView owner, FolderMergeBook book, IReadOnlyDictionary<MergeCellAddress, MergeChoice> choices,
        PreviewOptions options)
    {
        Owner = owner;
        Book = book;
        this.choices = choices;
        Dock = DockStyle.Fill;
        rows = book.Plan?.Changes.OrderByDescending(change => change.Disposition == MergeDisposition.Conflict).ToArray() ?? [];
        conflicts = rows.Where(row => row.Disposition == MergeDisposition.Conflict).ToArray();
        changesByCell = rows.ToDictionary(change => change.Cell);
        mergeChanges = rows.SelectMany(change => change.MergeRanges.Select(range => (range, change))).ToArray();
        notes = book.Notice ?? string.Join(Environment.NewLine,
            book.Plan?.Notices.Select(note => $"{note.Sheet}　{note.Reason}") ?? []);
        preview = new MergePreview(book, owner.UsesBase);
        comparison = new MergeComparisonView(book, owner.UsesBase);
        var layout = Appearance.Stack();
        var actions = Appearance.Toolbar();
        useProduction = Appearance.Button("取り込み1を採用", () => Choose(MergeChoice.Production));
        useIncoming = Appearance.Button("取り込み2を採用", () => Choose(MergeChoice.Incoming));
        reveal = Appearance.Button("プレビューで表示", () => OpenPreview(ChoiceTarget));
        nextUnresolved = Appearance.Button("次の未解決へ", MoveToUnresolved);
        actions.Controls.AddRange([useProduction, useIncoming, reveal, nextUnresolved, status]);
        preview.CellSelected += cell =>
        {
            previewChange = FindChange(cell);
            if (previewChange is { } change) preview.ShowChangeValues(change);
            UpdateActions();
        };
        comparison.DifferenceSelected += (sheet, difference, cell) =>
        {
            comparisonSheet = sheet;
            comparisonDifference = difference;
            var canAdopt = difference?.Kind is ChangeKind.Value or ChangeKind.Formula or ChangeKind.Added or ChangeKind.Removed ||
                difference is { Kind: ChangeKind.Merge, Parts: { } parts } &&
                parts.Any(part => part.Item is "値" or "数式" or "結合範囲");
            comparisonChange = canAdopt ? FindChange(cell) : null;
            UpdateActions();
        };
        Appearance.Row(layout, actions);
        Appearance.ConfigureGrid(grid);
        grid.VirtualMode = true;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AccessibleName = "ブックのセル差分とコンフリクト";
        grid.ZoomRequested += zoom.Apply;
        zoom.Changed += () => grid.SetZoom(zoom.Scale);
        foreach (var (title, weight) in new[] { ("シート", 15), ("セル・結合範囲", 13), ("共通ベース", 18),
                     ("取り込み1", 18), ("取り込み2", 18), ("処理", 24) })
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = title, FillWeight = weight,
                MinimumWidth = title == "セル・結合範囲" ? 110 : 64, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        grid.Columns[2].Visible = owner.UsesBase;
        sort = new GridSort(grid, SortChanges);
        grid.CellValueNeeded += (_, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= rows.Length) return;
            var change = rows[e.RowIndex];
            e.Value = e.ColumnIndex switch
            {
                0 => change.Cell.Sheet, 1 => change.Address, 2 => change.Base, 3 => change.Production, 4 => change.Incoming,
                _ => DecisionText(Decision(change))
            };
        };
        grid.CellToolTipTextNeeded += (_, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= rows.Length) return;
            var change = rows[e.RowIndex];
            e.ToolTipText = e.ColumnIndex == 1 ? change.Address :
                $"共通ベース　{change.Base}{Environment.NewLine}取り込み1　{change.Production}{Environment.NewLine}取り込み2　{change.Incoming}";
        };
        grid.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex >= 0 && e.RowIndex < rows.Length) OpenPreview(rows[e.RowIndex]);
        };
        grid.CurrentCellChanged += (_, _) => UpdateActions();
        grid.CellFormatting += (_, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= rows.Length || e.CellStyle is not { } style) return;
            var change = rows[e.RowIndex];
            switch (Decision(change))
            {
                case DecisionState.Unresolved:
                case DecisionState.Unsupported:
                    style.BackColor = Appearance.MergeConflict;
                    style.SelectionBackColor = Appearance.MergeConflictSelected;
                    break;
                case DecisionState.TakeIncoming when e.ColumnIndex == 5:
                    style.BackColor = Appearance.Added;
                    break;
            }
        };
        var cells = Appearance.Stack();
        Appearance.Row(cells, grid, fill: true);
        Appearance.Row(cells, details);
        Appearance.Row(cells, new ZoomBar(zoom));
        cellTab.Controls.Add(cells);
        sheetTab.Controls.Add(comparison);
        previewTab.Controls.Add(preview);
        views.TabPages.AddRange([sheetTab, cellTab, previewTab]);
        if (book.Plan is { } plan)
        {
            package = owner.UsesBase ? new PackageDiffView() : new PackageDiffView(book.ProductionPath, book.IncomingPath);
            var partTab = new TabPage($"ブック全体の差分 ({plan.PartChanges.Count + plan.PairParts.Count})");
            partTab.Controls.Add(package);
            views.TabPages.Add(partTab);
            if (owner.UsesBase) package.SetMerge(plan.PartChanges, book.BasePath!, book.ProductionPath!, book.IncomingPath!);
            else package.SetComparison(plan.PairParts, "取り込み1", "取り込み2");
        }
        views.SelectedTab = sheetTab;
        views.SelectedIndexChanged += (_, _) =>
        {
            if (views.SelectedTab == previewTab) preview.LoadPreview();
            if (views.SelectedTab == sheetTab) comparison.LoadComparison();
            UpdateActions();
        };
        Appearance.Row(layout, views, fill: true);
        Controls.Add(layout);
        SetPreviewOptions(options);
        grid.RowCount = rows.Length;
        if (grid.RowCount > 0) grid.CurrentCell = grid.Rows[0].Cells[0];
        VisibleChanged += (_, _) =>
        {
            if (!Visible) return;
            if (views.SelectedTab == sheetTab) comparison.LoadComparison();
            UpdateActions();
        };
        HandleCreated += (_, _) => { if (views.SelectedTab == sheetTab) comparison.LoadComparison(); };
        UpdateActions();
    }

    internal MergeView Owner { get; }
    internal FolderMergeBook Book { get; }
    private MergeCellChange? Current => grid.CurrentCell is { RowIndex: var index } && index >= 0 && index < rows.Length
        ? rows[index] : null;
    private MergeCellChange? ChoiceTarget => views.SelectedTab == previewTab ? previewChange :
        views.SelectedTab == cellTab ? Current : views.SelectedTab == sheetTab ? comparisonChange : null;

    private MergeCellChange? FindChange(MergeCellAddress? cell)
    {
        if (cell is null) return null;
        return changesByCell.GetValueOrDefault(cell) ?? mergeChanges.FirstOrDefault(item =>
            string.Equals(item.Change.Cell.Sheet, cell.Sheet, StringComparison.OrdinalIgnoreCase) &&
            item.Range.FirstRow <= cell.Row && item.Range.LastRow >= cell.Row && item.Range.ContainsColumn(cell.Column)).Change;
    }

    internal void SetPreviewOptions(PreviewOptions options)
    {
        preview.SetOptions(options);
        comparison.SetOptions(options);
        package?.SetThreeWay(options.ShowThreeWay);
    }

    internal void SetLocked(bool busy)
    {
        locked = busy;
        UpdateActions();
    }

    internal void OpenPreview(MergeCellChange? change = null)
    {
        if (change is not null) preview.Reveal(change.Cell);
        views.SelectedTab = previewTab;
        preview.LoadPreview();
        UpdateActions();
    }

    private void Choose(MergeChoice choice)
    {
        if (locked || ChoiceTarget is not { CanChoose: true } change) return;
        Owner.Choose(Book, change.Cell, choice);
        if (sort.Column == 5) SortChanges();
        else
        {
            grid.Invalidate();
            UpdateActions();
        }
    }

    private void MoveToUnresolved()
    {
        if (locked) return;
        var current = Array.IndexOf(conflicts, previewChange);
        for (var offset = 1; offset <= conflicts.Length; offset++)
        {
            var change = conflicts[(current + offset) % conflicts.Length];
            if (choices.ContainsKey(change.Cell)) continue;
            preview.Reveal(change.Cell);
            return;
        }
    }

    private void SortChanges()
    {
        var selected = Current;
        var column = grid.CurrentCell?.ColumnIndex ?? 0;
        var text = StringComparer.CurrentCultureIgnoreCase;
        rows = (sort.Column switch
        {
            0 => sort.Order(rows, row => row.Cell.Sheet, text),
            1 => sort.Order(rows, row => (row.Cell.Row, row.Cell.Column)),
            2 => sort.Order(rows, row => row.Base, text),
            3 => sort.Order(rows, row => row.Production, text),
            4 => sort.Order(rows, row => row.Incoming, text),
            _ => sort.Order(rows, Decision)
        }).ToArray();
        if (selected is not null) grid.CurrentCell = grid[column, Array.IndexOf(rows, selected)];
        grid.Invalidate();
        UpdateActions();
    }

    private void UpdateActions()
    {
        var inPreview = views.SelectedTab == previewTab;
        var unresolved = conflicts.Count(change => !choices.ContainsKey(change.Cell));
        useProduction.Enabled = useIncoming.Enabled = !locked && ChoiceTarget?.CanChoose == true;
        reveal.Visible = !inPreview;
        nextUnresolved.Visible = inPreview;
        nextUnresolved.Enabled = !locked && unresolved > 0;
        var selected = ChoiceTarget is { } pick ? $"　{pick.Address}　{DecisionText(Decision(pick))}" : "";
        status.Text = $"未解決のコンフリクト {unresolved} 件{selected}";
        reveal.Enabled = !locked && ChoiceTarget is not null;
        comparison.ShowMergeInfo(ComparisonInfo());
        var contents = Current is { } change ?
            (change.Reason is null ? "" : change.Reason + Environment.NewLine) +
            (Owner.UsesBase ? $"共通ベース　{change.Base}{Environment.NewLine}" : "") +
            $"取り込み1　{change.Production}{Environment.NewLine}取り込み2　{change.Incoming}" : "";
        details.Text = string.Join(Environment.NewLine, new[] { notes, contents }.Where(text => text.Length > 0));
        details.Visible = details.Text.Length > 0;
    }

    private string ComparisonInfo()
    {
        var notices = Book.Notice ?? string.Join(Environment.NewLine, Book.Plan?.Notices
            .Where(note => comparisonSheet is null || note.Sheet == "ブック" ||
                string.Equals(note.Sheet, comparisonSheet.Name, StringComparison.OrdinalIgnoreCase))
            .Select(note => $"{note.Sheet}　{note.Reason}") ?? []);
        var decision = comparisonChange is { } change
            ? $"{DecisionText(Decision(change))}　{change.Reason}" : "";
        if (comparisonChange is null && comparisonDifference is not null && notices.Length == 0)
            decision = comparisonDifference.Kind is ChangeKind.Added or ChangeKind.Removed
                ? "行・シート単位の差分です。セルごとの採用状態は「差分・コンフリクト」で確認できます。"
                : "この差分は取り込み1の内容を引き継ぎます。書式・行高・列幅や行・列の移動なども表示しています。";
        return string.Join(Environment.NewLine, new[] { notices, decision }.Where(text => !string.IsNullOrWhiteSpace(text)));
    }

    private enum DecisionState { Unresolved, Unsupported, TakeIncoming, AdoptedProduction, AdoptedIncoming, KeepProduction }

    private DecisionState Decision(MergeCellChange change)
    {
        if (change.Disposition == MergeDisposition.Unsupported) return DecisionState.Unsupported;
        if (choices.TryGetValue(change.Cell, out var choice))
            return choice == MergeChoice.Incoming ? DecisionState.AdoptedIncoming : DecisionState.AdoptedProduction;
        return change.Disposition switch
        {
            MergeDisposition.KeepProduction or MergeDisposition.AlreadyMatches => DecisionState.KeepProduction,
            MergeDisposition.TakeIncoming => DecisionState.TakeIncoming,
            _ => DecisionState.Unresolved
        };
    }

    private static string DecisionText(DecisionState state) => state switch
    {
        DecisionState.KeepProduction => "取り込み1を採用（自動）",
        DecisionState.TakeIncoming => "取り込み2を採用（自動）",
        DecisionState.Unsupported => "コンフリクト（手動解消）",
        DecisionState.AdoptedIncoming => "取り込み2を採用",
        DecisionState.AdoptedProduction => "取り込み1を採用",
        _ => "コンフリクト・未選択"
    };
}
