using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeView : UserControl
{
    private const int DecisionColumn = 5;
    private readonly CancellationTokenSource lifetime = new();
    private readonly TextBox basePath = new();
    private readonly TextBox productionPath = new();
    private readonly TextBox incomingPath = new();
    private readonly DataGridView grid = Appearance.Grid();
    private readonly Label summary = Appearance.Label("");
    private readonly TextBox notices = new()
    {
        Dock = DockStyle.Top, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical,
        Height = 62, Visible = false, BackColor = Appearance.Changed, BorderStyle = BorderStyle.FixedSingle,
        AccessibleName = "マージ対象外のシート"
    };
    private readonly Button analyze;
    private readonly Button save;
    private readonly Button useProduction;
    private readonly Button useIncoming;
    private readonly Dictionary<MergeCellAddress, MergeChoice> choices = [];
    private MergeCellChange[] rows = [];
    private WorkbookMergePlan? plan;
    private int conflictCount;
    private int automaticCount;
    private int unsupportedCount;
    private bool busy;

    internal MergeView()
    {
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var paths = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, BackColor = Appearance.Panel };
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        AddPath(paths, 0, "共通ベース", basePath);
        AddPath(paths, 1, "本番版", productionPath);
        AddPath(paths, 2, "取り込み元", incomingPath);
        Appearance.Row(layout, paths);

        var commands = Appearance.Toolbar();
        analyze = Appearance.Button("解析", async () => await Analyze(), primary: true);
        save = Appearance.Button("別名で保存…", async () => await Save());
        useProduction = Appearance.Button("本番版を採用", () => Choose(MergeChoice.Production));
        useIncoming = Appearance.Button("取り込み元を採用", () => Choose(MergeChoice.Incoming));
        commands.Controls.AddRange([analyze, useProduction, useIncoming, save]);
        Appearance.Row(layout, commands);
        Appearance.Row(layout, summary);
        Appearance.Row(layout, notices);

        grid.VirtualMode = true;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        foreach (var (label, weight) in new[] { ("シート", 15), ("セル", 7), ("共通ベース", 17),
                     ("本番版", 18), ("取り込み元", 18), ("処理", 23) })
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = label, FillWeight = weight, MinimumWidth = weight == 7 ? 48 : 80,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable
            });
        grid.CellValueNeeded += (_, e) =>
        {
            if (At(e.RowIndex) is not { } change) return;
            e.Value = e.ColumnIndex switch
            {
                0 => change.Cell.Sheet,
                1 => change.Cell.Address,
                2 => change.Base,
                3 => change.Production,
                4 => change.Incoming,
                _ => Decision(change)
            };
        };
        grid.CellToolTipTextNeeded += (_, e) =>
        {
            if (At(e.RowIndex) is not { } change) return;
            e.ToolTipText = e.ColumnIndex switch
            {
                2 => change.Base,
                3 => change.Production,
                4 => change.Incoming,
                DecisionColumn => change.Reason ?? Decision(change),
                _ => change.Cell.Sheet + " " + change.Cell.Address
            };
        };
        grid.CellFormatting += (_, e) =>
        {
            if (e.ColumnIndex != DecisionColumn || At(e.RowIndex) is not { } change || e.CellStyle is not { } style) return;
            style.BackColor = change.Disposition switch
            {
                MergeDisposition.Conflict when !choices.ContainsKey(change.Cell) => Appearance.Changed,
                MergeDisposition.TakeIncoming => Appearance.Added,
                MergeDisposition.Unsupported => Appearance.Removed,
                _ => Color.White
            };
        };
        grid.SelectionChanged += (_, _) => UpdateControls();
        Appearance.Row(layout, grid, fill: true);
        Controls.Add(layout);
        foreach (var box in new[] { basePath, productionPath, incomingPath })
            box.TextChanged += (_, _) =>
            {
                ResetPlan();
                summary.Text = "共通ベース・本番版・取り込み元を指定して解析します。";
                UpdateControls();
            };
        summary.Text = "共通ベース・本番版・取り込み元を指定して解析します。";
        UpdateControls();
    }

    private static void AddPath(TableLayoutPanel paths, int row, string label, TextBox box)
    {
        box.Dock = DockStyle.Fill;
        box.Margin = new Padding(5, 8, 5, 5);
        box.AccessibleName = label;
        paths.Controls.Add(Appearance.Label(label), 0, row);
        paths.Controls.Add(box, 1, row);
        paths.Controls.Add(Appearance.Button("参照…", () =>
        {
            using var dialog = new OpenFileDialog { Title = label, Filter = "Excel ブック (*.xlsx)|*.xlsx", CheckFileExists = true };
            if (File.Exists(box.Text)) dialog.FileName = box.Text;
            if (dialog.ShowDialog(paths) == DialogResult.OK) box.Text = dialog.FileName;
        }), 2, row);
    }

    private async Task Analyze()
    {
        var original = basePath.Text.Trim();
        var production = productionPath.Text.Trim();
        var incoming = incomingPath.Text.Trim();
        var token = lifetime.Token;
        ResetPlan();
        SetBusy(true, "3冊のセル内容を比較しています…");
        try
        {
            var result = await Task.Run(() => WorkbookMerger.Analyze(original, production, incoming, token), token);
            if (IsDisposed || token.IsCancellationRequested) return;
            plan = result;
            conflictCount = result.Changes.Count(change => change.Disposition == MergeDisposition.Conflict);
            automaticCount = result.Changes.Count(change => change.Disposition == MergeDisposition.TakeIncoming);
            unsupportedCount = result.Changes.Count(change => change.Disposition == MergeDisposition.Unsupported);
            rows = result.Changes.OrderByDescending(change => change.Disposition == MergeDisposition.Conflict).ToArray();
            notices.Text = string.Join(Environment.NewLine, result.Notices.Select(notice => $"{notice.Sheet}　{notice.Reason}"));
            notices.Visible = result.Notices.Count > 0;
            grid.RowCount = rows.Length;
            UpdateSummary();
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { if (!IsDisposed) summary.Text = error.Message; }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private async Task Save()
    {
        if (plan is not { } result) return;
        using var dialog = new SaveFileDialog
        {
            Title = "マージ結果を保存",
            Filter = "Excel ブック (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            FileName = Path.GetFileNameWithoutExtension(productionPath.Text.Trim()) + "-merged.xlsx"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var selection = new Dictionary<MergeCellAddress, MergeChoice>(choices);
        var token = lifetime.Token;
        SetBusy(true, "マージ結果を保存しています…");
        try
        {
            await Task.Run(() => result.Save(dialog.FileName, selection, token), token);
            if (!IsDisposed && !token.IsCancellationRequested) summary.Text = $"保存しました　{dialog.FileName}";
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { if (!IsDisposed) summary.Text = error.Message; }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private void Choose(MergeChoice choice)
    {
        if (grid.CurrentRow is not { } row || At(row.Index) is not { Disposition: MergeDisposition.Conflict } change) return;
        choices[change.Cell] = choice;
        grid.InvalidateCell(DecisionColumn, row.Index);
        UpdateSummary();
        UpdateControls();
    }

    private void UpdateSummary()
    {
        if (plan is null) return;
        summary.Text = plan.Changes.Count == 0 && plan.Notices.Count == 0
            ? "マージするセル差分はありません。"
            : $"自動取り込み {automaticCount} 件　未解決の競合 {conflictCount - choices.Count} / {conflictCount} 件" +
              $"　対象外 {unsupportedCount} セル・{plan.Notices.Count} シート";
    }

    private MergeCellChange? At(int row) => row >= 0 && row < rows.Length ? rows[row] : null;

    private string Decision(MergeCellChange change) => change.Disposition switch
    {
        MergeDisposition.KeepProduction => "本番版を保持",
        MergeDisposition.TakeIncoming => "自動取り込み",
        MergeDisposition.AlreadyMatches => "両側で同じ変更",
        MergeDisposition.Unsupported => "対象外（本番版を保持）",
        _ => !choices.TryGetValue(change.Cell, out var choice) ? "競合・未選択" :
            choice == MergeChoice.Incoming ? "取り込み元を採用" : "本番版を採用"
    };

    private void ResetPlan()
    {
        plan = null;
        rows = [];
        choices.Clear();
        grid.RowCount = 0;
        notices.Visible = false;
        conflictCount = automaticCount = unsupportedCount = 0;
    }

    private void SetBusy(bool value, string? message = null)
    {
        busy = value;
        Cursor = value ? Cursors.WaitCursor : Cursors.Default;
        if (message is not null) summary.Text = message;
        if (basePath.Parent is Control paths) paths.Enabled = !value;
        UpdateControls();
    }

    private void UpdateControls()
    {
        analyze.Enabled = !busy && !string.IsNullOrWhiteSpace(basePath.Text) &&
            !string.IsNullOrWhiteSpace(productionPath.Text) && !string.IsNullOrWhiteSpace(incomingPath.Text);
        save.Enabled = !busy && plan is not null && choices.Count == conflictCount;
        useProduction.Enabled = useIncoming.Enabled = !busy && grid.CurrentRow is { } row &&
            At(row.Index)?.Disposition == MergeDisposition.Conflict;
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
