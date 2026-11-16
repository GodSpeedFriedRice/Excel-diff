using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeBookView : UserControl
{
    private readonly DataGridView grid = new();
    private readonly GridSort sort;
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage cellTab = new("差分・コンフリクト");
    private readonly TabPage previewTab = new("全体プレビュー");
    private readonly Label status = Appearance.Label("");
    private readonly TextBox details = new()
    {
        Dock = DockStyle.Top, Height = 56, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
        AccessibleName = "選択した差分の詳細"
    };
    private readonly PackageDiffView package = new();
    private readonly MergePreview preview;
    private readonly Button useProduction;
    private readonly Button useIncoming;
    private readonly Button reveal;
    private MergeCellChange[] rows;
    private readonly IReadOnlyDictionary<MergeCellAddress, MergeChoice> choices;
    private readonly int conflicts;
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
        conflicts = rows.Count(row => row.Disposition == MergeDisposition.Conflict);
        notes = book.Notice ?? string.Join(Environment.NewLine,
            book.Plan?.Notices.Select(note => $"{note.Sheet}　{note.Reason}") ?? []);
        preview = new MergePreview(book);
        var layout = Appearance.Stack();
        var actions = Appearance.Toolbar();
        useProduction = Appearance.Button("取り込み1を採用", () => Choose(MergeChoice.Production));
        useIncoming = Appearance.Button("取り込み2を採用", () => Choose(MergeChoice.Incoming));
        reveal = Appearance.Button("プレビューで表示", () => OpenPreview(Current));
        actions.Controls.AddRange([useProduction, useIncoming, reveal, status]);
        Appearance.Row(layout, actions);
        Appearance.ConfigureGrid(grid);
        grid.VirtualMode = true;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AccessibleName = "ブックのセル差分とコンフリクト";
        foreach (var (title, weight) in new[] { ("シート", 15), ("セル", 7), ("共通ベース", 18),
                     ("取り込み1", 18), ("取り込み2", 18), ("処理", 24) })
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = title, FillWeight = weight,
                MinimumWidth = weight == 7 ? 46 : 64, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        sort = new GridSort(grid, SortChanges);
        grid.CellValueNeeded += (_, e) =>
        {
            if (e.RowIndex < 0 || e.RowIndex >= rows.Length) return;
            var change = rows[e.RowIndex];
            e.Value = e.ColumnIndex switch
            {
                0 => change.Cell.Sheet, 1 => change.Cell.Address, 2 => change.Base, 3 => change.Production, 4 => change.Incoming,
                _ => DecisionText(Decision(change))
            };
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
        cellTab.Controls.Add(cells);
        previewTab.Controls.Add(preview);
        views.TabPages.AddRange([cellTab, previewTab]);
        if (book.Plan is { } plan)
        {
            var partTab = new TabPage($"ブック全体の差分 ({plan.PartChanges.Count})");
            partTab.Controls.Add(package);
            views.TabPages.Add(partTab);
            package.SetMerge(plan.PartChanges, book.BasePath!, book.ProductionPath!, book.IncomingPath!);
        }
        else package.Dispose();
        views.SelectedTab = cellTab;
        views.SelectedIndexChanged += (_, _) =>
        {
            if (views.SelectedTab == previewTab) preview.LoadPreview();
            UpdateActions();
        };
        Appearance.Row(layout, views, fill: true);
        Controls.Add(layout);
        SetPreviewOptions(options);
        grid.RowCount = rows.Length;
        if (grid.RowCount > 0) grid.CurrentCell = grid.Rows[0].Cells[0];
        VisibleChanged += (_, _) => { if (Visible) UpdateActions(); };
        UpdateActions();
    }

    internal MergeView Owner { get; }
    internal FolderMergeBook Book { get; }
    private MergeCellChange? Current => grid.CurrentCell is { RowIndex: var index } && index >= 0 && index < rows.Length
        ? rows[index] : null;

    internal void SetPreviewOptions(PreviewOptions options)
    {
        preview.SetOptions(options);
        if (!package.IsDisposed) package.SetThreeWay(options.ShowThreeWay);
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
    }

    private void Choose(MergeChoice choice)
    {
        if (locked || Current is not { Disposition: MergeDisposition.Conflict } change) return;
        Owner.Choose(Book, change.Cell, choice);
        if (sort.Column == 5) SortChanges();
        else
        {
            grid.InvalidateRow(grid.CurrentCell!.RowIndex);
            UpdateActions();
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
        useProduction.Enabled = useIncoming.Enabled = !locked && views.SelectedTab == cellTab &&
            Current?.Disposition == MergeDisposition.Conflict;
        reveal.Enabled = views.SelectedTab != previewTab;
        status.Text = $"未解決のコンフリクト {conflicts - choices.Count} 件";
        var contents = Current is { } change ? change.Reason ??
            $"共通ベース　{change.Base}{Environment.NewLine}取り込み1　{change.Production}{Environment.NewLine}取り込み2　{change.Incoming}" : "";
        details.Text = string.Join(Environment.NewLine, new[] { notes, contents }.Where(text => text.Length > 0));
        details.Visible = details.Text.Length > 0;
    }

    private enum DecisionState { Unresolved, Unsupported, TakeIncoming, AdoptedProduction, AdoptedIncoming, KeepProduction, AlreadyMatches }

    private DecisionState Decision(MergeCellChange change) => change.Disposition switch
    {
        MergeDisposition.KeepProduction => DecisionState.KeepProduction,
        MergeDisposition.TakeIncoming => DecisionState.TakeIncoming,
        MergeDisposition.AlreadyMatches => DecisionState.AlreadyMatches,
        MergeDisposition.Unsupported => DecisionState.Unsupported,
        _ => !choices.TryGetValue(change.Cell, out var choice) ? DecisionState.Unresolved :
            choice == MergeChoice.Incoming ? DecisionState.AdoptedIncoming : DecisionState.AdoptedProduction
    };

    private static string DecisionText(DecisionState state) => state switch
    {
        DecisionState.KeepProduction => "取り込み1を保持",
        DecisionState.TakeIncoming => "自動取り込み",
        DecisionState.AlreadyMatches => "両側で同じ変更",
        DecisionState.Unsupported => "コンフリクト（手動解消）",
        DecisionState.AdoptedIncoming => "取り込み2を採用",
        DecisionState.AdoptedProduction => "取り込み1を採用",
        _ => "コンフリクト・未選択"
    };
}
