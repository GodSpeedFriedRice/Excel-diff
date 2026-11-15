using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeView : UserControl
{
    private readonly CancellationTokenSource lifetime = new();
    private readonly CheckBox folders = new() { Text = "フォルダを指定", Checked = true, AutoSize = true };
    private readonly TextBox basePath = new();
    private readonly TextBox productionPath = new();
    private readonly TextBox incomingPath = new();
    private readonly DataGridView fileList = Appearance.Grid();
    private readonly Label summary = Appearance.Label("");
    private readonly Label notice = Appearance.Label("");
    private readonly ComboBox filter = new() { Width = 114, DropDownStyle = ComboBoxStyle.DropDownList,
        AccessibleName = "ファイルの絞り込み" };
    private readonly Button analyze;
    private readonly Button changeInputs;
    private readonly Button save;
    private readonly Button open;
    private readonly Button preview;
    private readonly TableLayoutPanel paths;
    private readonly string recentTargetsPath;
    private readonly Dictionary<string, Dictionary<MergeCellAddress, MergeChoice>> choices = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, int> conflicts = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> review = new(StringComparer.OrdinalIgnoreCase);
    private IReadOnlyList<FolderMergeBook> books = [];
    private FolderMergeBook[] visible = [];
    private FolderMergePlan? folderPlan;
    private int unresolved;
    private bool busy;
    private bool rendering;
    private bool inputsVisible = true;

    internal event Action<MergeView, FolderMergeBook, bool>? OpenRequested;
    internal event Action<MergeView>? ResultsCleared;
    internal event Action<bool>? BusyChanged;

    internal MergeView() : this(LocalSettings.PathFor("last-merge.json")) { }

    internal MergeView(string recentTargetsPath)
    {
        this.recentTargetsPath = recentTargetsPath;
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        paths = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, BackColor = Appearance.Panel };
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        folders.Margin = new Padding(6, 5, 6, 4);
        paths.Controls.Add(folders, 1, 0);
        AddPath(1, "共通ベース", basePath);
        AddPath(2, "本番版", productionPath);
        AddPath(3, "取り込み元", incomingPath);
        Appearance.Row(layout, paths);

        var commands = Appearance.Toolbar();
        analyze = Appearance.Button("解析", async () => await Analyze(), primary: true);
        changeInputs = Appearance.Button("比較対象を変更", () => ShowInputs(!inputsVisible));
        save = Appearance.Button("結果フォルダに保存…", async () => await Save());
        open = Appearance.Button("詳細を開く", () => OpenCurrent(false));
        preview = Appearance.Button("プレビュー", () => OpenCurrent(true));
        filter.Items.AddRange(["すべて", "未解決の競合", "手動確認"]);
        filter.SelectedIndex = 0;
        filter.Margin = new Padding(6, 7, 0, 4);
        filter.SelectedIndexChanged += (_, _) => RenderFiles();
        commands.Controls.AddRange([analyze, changeInputs, filter, open, preview, save]);
        Appearance.Row(layout, commands);
        Appearance.Row(layout, summary);

        fileList.VirtualMode = true;
        fileList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        fileList.AccessibleName = "3方向マージのファイル一覧";
        foreach (var (title, weight) in new[] { ("ファイル（相対パス）", 50), ("状態", 30), ("未解決の競合", 20) })
            fileList.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = title, FillWeight = weight, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        fileList.CellValueNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is not { } book) return;
            e.Value = e.ColumnIndex switch { 0 => book.RelativePath, 1 => BookStatus(book),
                _ => book.Plan is null ? "—" : Pending(book).ToString() };
        };
        fileList.CellToolTipTextNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is { } book) e.ToolTipText = e.ColumnIndex == 0 ? book.RelativePath : book.Notice;
        };
        fileList.CurrentCellChanged += (_, _) => { if (!rendering) UpdateControls(); };
        fileList.CellDoubleClick += (_, e) => { if (e.RowIndex >= 0) OpenCurrent(false); };
        fileList.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = e.SuppressKeyPress = true;
            OpenCurrent(false);
        };
        Appearance.Row(layout, fileList, fill: true);
        Appearance.Row(layout, notice);
        Controls.Add(layout);
        if (LocalSettings.Load<MergeTargets>(recentTargetsPath) is { } previous &&
            !string.IsNullOrWhiteSpace(previous.Base) && !string.IsNullOrWhiteSpace(previous.Production) &&
            !string.IsNullOrWhiteSpace(previous.Incoming))
        {
            basePath.Text = previous.Base;
            productionPath.Text = previous.Production;
            incomingPath.Text = previous.Incoming;
            folders.Checked = previous.Folders;
        }
        foreach (var box in new[] { basePath, productionPath, incomingPath })
            box.TextChanged += (_, _) => ResetInputs();
        folders.CheckedChanged += (_, _) => ResetInputs();
        ResetInputs();
    }

    internal bool Busy => busy;

    internal IReadOnlyDictionary<MergeCellAddress, MergeChoice> ChoicesFor(FolderMergeBook book)
    {
        if (!choices.TryGetValue(book.RelativePath, out var selected)) choices.Add(book.RelativePath, selected = []);
        return selected;
    }

    internal void Choose(FolderMergeBook book, MergeCellAddress cell, MergeChoice choice)
    {
        if (busy || !choices.TryGetValue(book.RelativePath, out var selected)) return;
        if (!selected.ContainsKey(cell)) unresolved--;
        selected[cell] = choice;
        RenderFiles();
        UpdateSummary();
    }

    private FolderMergeBook? Current => FileAt(fileList.CurrentCell?.RowIndex ?? -1);
    private FolderMergeBook? FileAt(int index) => index >= 0 && index < visible.Length ? visible[index] : null;
    private int Pending(FolderMergeBook book) => conflicts[book.RelativePath] - (choices.GetValueOrDefault(book.RelativePath)?.Count ?? 0);

    private void OpenCurrent(bool showPreview)
    {
        if (!busy && Current is { } book) OpenRequested?.Invoke(this, book, showPreview);
    }

    private void AddPath(int row, string label, TextBox box)
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
                if (folder.ShowDialog(this) == DialogResult.OK) box.Text = folder.SelectedPath;
            }
            else
            {
                using var dialog = new OpenFileDialog { Title = label, Filter = "Excel ブック (*.xlsx)|*.xlsx", CheckFileExists = true };
                if (File.Exists(box.Text)) dialog.FileName = box.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK) box.Text = dialog.FileName;
            }
        }), 2, row);
    }

    private async Task Analyze()
    {
        var compareFolders = folders.Checked;
        var token = lifetime.Token;
        ResetPlan();
        SetBusy(true, compareFolders ? "3つのフォルダを比較しています…" : "3冊のセル内容を比較しています…");
        try
        {
            var original = Path.GetFullPath(ReadPath(basePath));
            var production = Path.GetFullPath(ReadPath(productionPath));
            var incoming = Path.GetFullPath(ReadPath(incomingPath));
            if (compareFolders)
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
                LoadBooks(result.Books);
            }
            else
            {
                var result = await Task.Run(() => WorkbookMerger.Analyze(original, production, incoming, token), token);
                if (IsDisposed || token.IsCancellationRequested) return;
                LoadBooks([new(Path.GetFileName(production), original, production, incoming, result, null)]);
            }
            LocalSettings.Save(recentTargetsPath, new MergeTargets(original, production, incoming, compareFolders));
            ShowInputs(false);
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error)) { if (!IsDisposed) summary.Text = error.Message; }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private sealed record MergeTargets(string Base, string Production, string Incoming, bool Folders);

    private void LoadBooks(IReadOnlyList<FolderMergeBook> result)
    {
        foreach (var book in result)
        {
            var total = book.Plan?.Changes.Count(change => change.Disposition == MergeDisposition.Conflict) ?? 0;
            conflicts.Add(book.RelativePath, total);
            unresolved += total;
            if (book.Plan is null || book.Plan.Notices.Count > 0 ||
                book.Plan.Changes.Any(change => change.Disposition == MergeDisposition.Unsupported))
                review.Add(book.RelativePath);
        }
        books = result.OrderByDescending(book => conflicts[book.RelativePath] > 0)
            .ThenByDescending(book => review.Contains(book.RelativePath)).ToArray();
        RenderFiles();
        UpdateSummary();
    }

    private void RenderFiles()
    {
        var old = Current;
        visible = books.Where(book => filter.SelectedIndex switch
        {
            1 => Pending(book) > 0,
            2 => review.Contains(book.RelativePath),
            _ => true
        }).ToArray();
        rendering = true;
        try
        {
            fileList.RowCount = visible.Length;
            var index = old is null ? -1 : Array.IndexOf(visible, old);
            if (visible.Length > 0) fileList.CurrentCell = fileList.Rows[Math.Max(0, index)].Cells[0];
            fileList.Invalidate();
        }
        finally { rendering = false; }
        UpdateControls();
    }

    private async Task Save()
    {
        if (folderPlan is { } batch)
        {
            using var folder = new FolderBrowserDialog { Description = "マージ結果を作成する親フォルダ", UseDescriptionForTitle = true };
            if (folder.ShowDialog(this) != DialogResult.OK) return;
            var name = Path.GetFileName(Path.TrimEndingDirectorySeparator(Path.GetFullPath(ReadPath(productionPath))));
            var output = Path.Combine(folder.SelectedPath, (name.Length > 0 ? name : "ExcelDiff") + "-merged");
            var selections = choices.ToDictionary(item => item.Key,
                item => (IReadOnlyDictionary<MergeCellAddress, MergeChoice>)new Dictionary<MergeCellAddress, MergeChoice>(item.Value),
                StringComparer.OrdinalIgnoreCase);
            await SaveTo(output, (path, token) => batch.Save(path, selections, token));
            return;
        }
        if (books.FirstOrDefault() is not { Plan: { } plan } book) return;
        using var dialog = new SaveFileDialog
        {
            Title = "マージ結果を保存", Filter = "Excel ブック (*.xlsx)|*.xlsx", DefaultExt = "xlsx",
            FileName = Path.GetFileNameWithoutExtension(book.RelativePath) + "-merged.xlsx"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var selection = new Dictionary<MergeCellAddress, MergeChoice>(ChoicesFor(book));
        await SaveTo(dialog.FileName, (path, token) => plan.Save(path, selection, token));
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
        catch (Exception error) when (FolderComparer.IsReadError(error) || error is InvalidOperationException)
        {
            if (!IsDisposed) summary.Text = error.Message;
        }
        finally { if (!IsDisposed) SetBusy(false); }
    }

    private void UpdateSummary() =>
        summary.Text = $"{books.Count} ファイル　未解決の競合 {unresolved} セル　手動確認 {review.Count} ファイル";

    private string BookStatus(FolderMergeBook book)
    {
        if (Pending(book) > 0) return "競合あり";
        if (review.Contains(book.RelativePath)) return "手動確認";
        if (conflicts[book.RelativePath] > 0) return "競合を解決済み";
        if (book.Plan!.Changes.Any(change => change.Disposition == MergeDisposition.TakeIncoming)) return "自動取り込みあり";
        return book.Plan.PartChanges.Count > 0 ? "差分あり" : "同一";
    }

    private void ResetPlan()
    {
        ResultsCleared?.Invoke(this);
        books = [];
        folderPlan = null;
        choices.Clear();
        conflicts.Clear();
        review.Clear();
        unresolved = 0;
        RenderFiles();
    }

    private static string ReadPath(TextBox box) => box.Text.Trim().Trim('"');

    private void ResetInputs()
    {
        ResetPlan();
        save.Text = folders.Checked ? "結果フォルダに保存…" : "別名で保存…";
        summary.Text = folders.Checked ? "共通ベース・本番版・取り込み元のフォルダを指定します。" :
            "共通ベース・本番版・取り込み元の Excel ファイルを指定します。";
        ShowInputs(true);
    }

    private void ShowInputs(bool show)
    {
        paths.Visible = inputsVisible = show;
        changeInputs.Text = show ? "比較対象を閉じる" : "比較対象を変更";
        UpdateControls();
    }

    private void SetBusy(bool value, string? message = null)
    {
        busy = value;
        Cursor = value ? Cursors.WaitCursor : Cursors.Default;
        if (message is not null) summary.Text = message;
        paths.Enabled = fileList.Enabled = !value;
        BusyChanged?.Invoke(value);
        UpdateControls();
    }

    private void UpdateControls()
    {
        analyze.Visible = inputsVisible;
        analyze.Enabled = !busy && ReadPath(basePath).Length > 0 && ReadPath(productionPath).Length > 0 && ReadPath(incomingPath).Length > 0;
        changeInputs.Visible = folderPlan is not null || books.Count > 0;
        changeInputs.Enabled = !busy;
        filter.Enabled = !busy && books.Count > 0;
        open.Enabled = preview.Enabled = !busy && Current is not null;
        save.Enabled = !busy && unresolved == 0 && books.Any(book => book.ProductionPath is not null);
        notice.Text = Current is { } book ? book.Notice ??
            (review.Contains(book.RelativePath) ? "詳細でマージ対象外の項目を確認できます。" : "") : "";
        notice.Visible = notice.Text.Length > 0;
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
