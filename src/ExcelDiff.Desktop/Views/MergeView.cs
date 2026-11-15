using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeView : UserControl
{
    private const int DecisionColumn = 5;
    private readonly CancellationTokenSource lifetime = new();
    private readonly CheckBox folders = new() { Text = "フォルダを指定", Checked = true, AutoSize = true };
    private readonly TextBox basePath = new();
    private readonly TextBox productionPath = new();
    private readonly TextBox incomingPath = new();
    private readonly DataGridView grid = Appearance.Grid();
    private readonly DataGridView fileList = Appearance.Grid();
    private readonly TabControl views = new() { Dock = DockStyle.Fill };
    private readonly TabPage cellTab = new("セル差分");
    private readonly TabPage packageTab = new("ブック全体の差分");
    private readonly PackageDiffView packageDiff = new();
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
    private readonly Dictionary<string, Dictionary<MergeCellAddress, MergeChoice>> fileChoices =
        new(StringComparer.OrdinalIgnoreCase);
    private Dictionary<MergeCellAddress, MergeChoice> choices = [];
    private MergeCellChange[] rows = [];
    private WorkbookMergePlan? plan;
    private FolderMergePlan? folderPlan;
    private FolderMergeBook? selectedBook;
    private Dictionary<string, int> fileConflicts = [];
    private int unresolved;
    private int reviewFiles;
    private bool hasProductionFiles;
    private int conflictCount;
    private int automaticCount;
    private int unsupportedCount;
    private bool busy;

    internal MergeView()
    {
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var mode = Appearance.Toolbar();
        folders.Margin = new Padding(6, 5, 6, 4);
        mode.Controls.Add(folders);
        Appearance.Row(layout, mode);
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
        fileList.Dock = DockStyle.Top;
        fileList.Height = 121;
        fileList.VirtualMode = true;
        fileList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        fileList.AccessibleName = "3方向マージの対象ファイル";
        foreach (var (title, weight) in new[] { ("ファイル（相対パス）", 50), ("解析結果", 35), ("未解決の競合", 15) })
            fileList.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = title, FillWeight = weight, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        fileList.CellValueNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is not { } book) return;
            e.Value = e.ColumnIndex switch
            {
                0 => book.RelativePath,
                1 => BookStatus(book),
                _ => book.Plan is null ? "—" :
                    (fileConflicts[book.RelativePath] - (fileChoices.GetValueOrDefault(book.RelativePath)?.Count ?? 0)).ToString()
            };
        };
        fileList.CellToolTipTextNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is { } book) e.ToolTipText = e.ColumnIndex == 0 ? book.RelativePath : book.Notice;
        };
        fileList.CurrentCellChanged += (_, _) => SelectBook(FileAt(fileList.CurrentCell?.RowIndex ?? -1));
        fileList.Visible = false;
        Appearance.Row(layout, fileList);
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
        cellTab.Controls.Add(grid);
        packageTab.Controls.Add(packageDiff);
        views.TabPages.AddRange([cellTab, packageTab]);
        views.SelectedIndexChanged += (_, _) => UpdateControls();
        Appearance.Row(layout, views, fill: true);
        Controls.Add(layout);
        foreach (var box in new[] { basePath, productionPath, incomingPath })
            box.TextChanged += (_, _) => ResetInputs();
        folders.CheckedChanged += (_, _) => ResetInputs();
        ResetInputs();
    }

    private void AddPath(TableLayoutPanel paths, int row, string label, TextBox box)
    {
        box.Dock = DockStyle.Fill;
        box.Margin = new Padding(5, 8, 5, 5);
        box.AccessibleName = label;
        paths.Controls.Add(Appearance.Label(label), 0, row);
        paths.Controls.Add(box, 1, row);
        paths.Controls.Add(Appearance.Button("参照…", () =>
        {
            if (folders.Checked)
            {
                using var folder = new FolderBrowserDialog { Description = label, UseDescriptionForTitle = true, SelectedPath = box.Text };
                if (folder.ShowDialog(paths) == DialogResult.OK) box.Text = folder.SelectedPath;
            }
            else
            {
                using var dialog = new OpenFileDialog { Title = label, Filter = "Excel ブック (*.xlsx)|*.xlsx", CheckFileExists = true };
                if (File.Exists(box.Text)) dialog.FileName = box.Text;
                if (dialog.ShowDialog(paths) == DialogResult.OK) box.Text = dialog.FileName;
            }
        }), 2, row);
    }

    private async Task Analyze()
    {
        var original = ReadPath(basePath);
        var production = ReadPath(productionPath);
        var incoming = ReadPath(incomingPath);
        var token = lifetime.Token;
        ResetPlan();
        SetBusy(true, folders.Checked ? "3つのフォルダを比較しています…" : "3冊のセル内容を比較しています…");
        try
        {
            if (folders.Checked)
            {
                FolderComparisonProgress? progress = null;
                using var timer = new System.Windows.Forms.Timer { Interval = 150 };
                timer.Tick += (_, _) =>
                {
                    if (Volatile.Read(ref progress) is { } count)
                        summary.Text = $"解析中　{count.Completed:N0} / {count.Total:N0} ファイル";
                };
                timer.Start();
                var result = await Task.Run(() => FolderMerger.Analyze(original, production, incoming, token,
                    value => Volatile.Write(ref progress, value)), token);
                if (IsDisposed || token.IsCancellationRequested) return;
                folderPlan = result;
                fileConflicts = result.Books.ToDictionary(book => book.RelativePath,
                    book => book.Plan?.Changes.Count(change => change.Disposition == MergeDisposition.Conflict) ?? 0,
                    StringComparer.OrdinalIgnoreCase);
                unresolved = fileConflicts.Values.Sum();
                reviewFiles = result.Books.Count(NeedsReview);
                hasProductionFiles = result.Books.Any(book => book.ProductionPath is not null);
                fileList.Visible = true;
                fileList.RowCount = result.Books.Count;
                if (fileList.RowCount > 0) fileList.CurrentCell = fileList.Rows[0].Cells[0];
                SelectBook(FileAt(0));
            }
            else
            {
                var result = await Task.Run(() => WorkbookMerger.Analyze(original, production, incoming, token), token);
                if (IsDisposed || token.IsCancellationRequested) return;
                ShowBook(result, original, production, incoming);
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { if (!IsDisposed) summary.Text = error.Message; }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private async Task Save()
    {
        if (folderPlan is { } batch)
        {
            using var folder = new FolderBrowserDialog { Description = "マージ結果を作成する親フォルダ", UseDescriptionForTitle = true };
            if (folder.ShowDialog(this) != DialogResult.OK) return;
            var name = Path.GetFileName(Path.TrimEndingDirectorySeparator(Path.GetFullPath(ReadPath(productionPath))));
            var output = Path.Combine(folder.SelectedPath, (name.Length > 0 ? name : "ExcelDiff") + "-merged");
            var selections = fileChoices.ToDictionary(item => item.Key,
                item => (IReadOnlyDictionary<MergeCellAddress, MergeChoice>)item.Value, StringComparer.OrdinalIgnoreCase);
            await SaveTo(output, (path, token) => batch.Save(path, selections, token));
            return;
        }
        if (plan is not { } result) return;
        using var dialog = new SaveFileDialog
        {
            Title = "マージ結果を保存",
            Filter = "Excel ブック (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            FileName = Path.GetFileNameWithoutExtension(productionPath.Text.Trim()) + "-merged.xlsx"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        await SaveTo(dialog.FileName, (path, token) => result.Save(path, choices, token));
    }

    private async Task SaveTo(string output, Action<string, CancellationToken> saveResult)
    {
        var token = lifetime.Token;
        SetBusy(true, "マージ結果を保存しています…");
        try
        {
            await Task.Run(() => saveResult(output, token), token);
            if (!IsDisposed && !token.IsCancellationRequested) summary.Text = $"保存しました　{output}";
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) { if (!IsDisposed) summary.Text = error.Message; }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private void Choose(MergeChoice choice)
    {
        if (grid.CurrentRow is not { } row || At(row.Index) is not { Disposition: MergeDisposition.Conflict } change) return;
        if (selectedBook is not null && !choices.ContainsKey(change.Cell)) unresolved--;
        choices[change.Cell] = choice;
        if (selectedBook is not null && fileList.CurrentCell is { } file) fileList.InvalidateRow(file.RowIndex);
        grid.InvalidateCell(DecisionColumn, row.Index);
        UpdateSummary();
        UpdateControls();
    }

    private void UpdateSummary()
    {
        if (folderPlan is { } batch)
        {
            summary.Text = $"{batch.Books.Count} ファイル　未解決の競合 {unresolved} セル" +
                $"　手動確認 {reviewFiles} ファイル　｜　保存に反映するのはセル差分です。";
            return;
        }
        if (plan is null) return;
        summary.Text = plan.Changes.Count == 0 && plan.Notices.Count == 0
            ? "マージ対象のセル差分なし"
            : $"自動取り込み {automaticCount} 件　未解決の競合 {conflictCount - choices.Count} / {conflictCount} 件" +
              $"　対象外 {unsupportedCount} セル・{plan.Notices.Count} シート";
        if (plan.PartChanges.Count > 0) summary.Text += "　｜　保存に反映するのはセル差分の採用結果です。";
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
        folderPlan = null;
        selectedBook = null;
        fileChoices.Clear();
        fileConflicts.Clear();
        unresolved = 0;
        reviewFiles = 0;
        hasProductionFiles = false;
        fileList.RowCount = 0;
        fileList.Visible = false;
        choices = [];
        ShowBook(null, null, null, null);
    }

    private void ShowBook(WorkbookMergePlan? result, string? original, string? production, string? incoming, string? notice = null)
    {
        grid.RowCount = 0;
        plan = result;
        rows = result?.Changes.OrderByDescending(change => change.Disposition == MergeDisposition.Conflict).ToArray() ?? [];
        conflictCount = result?.Changes.Count(change => change.Disposition == MergeDisposition.Conflict) ?? 0;
        automaticCount = result?.Changes.Count(change => change.Disposition == MergeDisposition.TakeIncoming) ?? 0;
        unsupportedCount = result?.Changes.Count(change => change.Disposition == MergeDisposition.Unsupported) ?? 0;
        if (result is null)
        {
            packageDiff.ResetMerge();
            packageTab.Text = "ブック全体の差分";
        }
        else
        {
            packageDiff.SetMerge(result.PartChanges, original!, production!, incoming!);
            packageTab.Text = $"ブック全体の差分 ({result.PartChanges.Count})";
        }
        notices.Text = notice ?? (result is null ? "" :
            string.Join(Environment.NewLine, result.Notices.Select(item => $"{item.Sheet}　{item.Reason}")));
        notices.Visible = notices.Text.Length > 0;
        grid.RowCount = rows.Length;
        UpdateSummary();
        UpdateControls();
    }

    private void SelectBook(FolderMergeBook? book)
    {
        if (folderPlan is null) return;
        selectedBook = book;
        choices = [];
        if (book?.Plan is not null)
        {
            if (fileChoices.TryGetValue(book.RelativePath, out var saved)) choices = saved;
            else fileChoices.Add(book.RelativePath, choices);
        }
        ShowBook(book?.Plan, book?.BasePath, book?.ProductionPath, book?.IncomingPath, book?.Notice);
    }

    private FolderMergeBook? FileAt(int index) =>
        folderPlan is { } batch && index >= 0 && index < batch.Books.Count ? batch.Books[index] : null;

    private static bool NeedsReview(FolderMergeBook book) => book.Plan is null || book.Plan.Notices.Count > 0 ||
        book.Plan.Changes.Any(change => change.Disposition == MergeDisposition.Unsupported);

    private static string BookStatus(FolderMergeBook book) => NeedsReview(book) ? "手動確認" :
        book.Plan!.PartChanges.Count > 0 ? "差分あり" : "同一";

    private static string ReadPath(TextBox box) => box.Text.Trim().Trim('"');

    private void ResetInputs()
    {
        ResetPlan();
        save.Text = folders.Checked ? "結果フォルダに保存…" : "別名で保存…";
        summary.Text = folders.Checked ? "共通ベース・本番版・取り込み元のフォルダを指定します。" :
            "共通ベース・本番版・取り込み元の Excel ファイルを指定します。";
    }

    private void SetBusy(bool value, string? message = null)
    {
        busy = value;
        Cursor = value ? Cursors.WaitCursor : Cursors.Default;
        if (message is not null) summary.Text = message;
        if (basePath.Parent is Control paths) paths.Enabled = !value;
        folders.Enabled = !value;
        fileList.Enabled = !value;
        UpdateControls();
    }

    private void UpdateControls()
    {
        analyze.Enabled = !busy && !string.IsNullOrWhiteSpace(basePath.Text) &&
            !string.IsNullOrWhiteSpace(productionPath.Text) && !string.IsNullOrWhiteSpace(incomingPath.Text);
        save.Enabled = !busy && (folderPlan is not null ? unresolved == 0 && hasProductionFiles :
            plan is not null && choices.Count == conflictCount);
        useProduction.Enabled = useIncoming.Enabled = !busy && views.SelectedTab == cellTab && grid.CurrentRow is { } row &&
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
