using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class MergeView : UserControl
{
    private readonly CancellationTokenSource lifetime = new();
    private readonly CheckBox folders = new() { Text = "フォルダを指定", Checked = true, AutoSize = true };
    private readonly CheckBox useBase = new() { Text = "共通ベースを使う", Checked = true, AutoSize = true };
    private readonly TextBox basePath = new();
    private readonly TextBox productionPath = new();
    private readonly TextBox incomingPath = new();
    private readonly DataGridView fileList = Appearance.Grid();
    private readonly GridSort fileSort;
    private readonly Label summary = Appearance.Label("");
    private readonly Label notice = Appearance.Label("");
    private readonly ComboBox filter = new() { Width = 144, DropDownStyle = ComboBoxStyle.DropDownList,
        AccessibleName = "ファイルの絞り込み" };
    private readonly Button analyze;
    private readonly Button changeInputs;
    private readonly Button save;
    private readonly Button open;
    private readonly Button preview;
    private readonly TableLayoutPanel paths;
    private readonly string recentTargetsPath;
    private readonly Dictionary<string, Dictionary<MergeCellAddress, MergeChoice>> choices = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, HashSet<MergeCellAddress>> conflicts = new(StringComparer.OrdinalIgnoreCase);
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
    internal event Action? ProgressChanged;

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
        var modes = Appearance.Toolbar();
        modes.Controls.AddRange([folders, useBase]);
        paths.Controls.Add(modes, 0, 0);
        paths.SetColumnSpan(modes, 3);
        AddPath(1, "共通ベース", basePath);
        AddPath(2, "取り込み1", productionPath);
        AddPath(3, "取り込み2", incomingPath);
        Appearance.Row(layout, paths);

        var commands = Appearance.Toolbar();
        analyze = Appearance.Button("比較", async () => await Analyze(), primary: true);
        changeInputs = Appearance.Button("比較対象を変更", () => ShowInputs(!inputsVisible));
        save = Appearance.Button("結果フォルダに保存…", async () => await Save());
        open = Appearance.Button("詳細を開く", () => OpenCurrent(false));
        preview = Appearance.Button("プレビュー", () => OpenCurrent(true));
        filter.Items.AddRange(["すべて", "コンフリクト", "手動解消が必要"]);
        filter.SelectedIndex = 0;
        filter.Margin = new Padding(6, 7, 0, 4);
        filter.SelectedIndexChanged += (_, _) => RenderFiles();
        commands.Controls.AddRange([analyze, changeInputs, filter, open, preview, save]);
        Appearance.Row(layout, commands);
        Appearance.Row(layout, summary);

        fileList.VirtualMode = true;
        fileList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        fileList.AccessibleName = "マージ比較のファイル一覧";
        foreach (var (title, weight) in new[] { ("ファイル名", 29), ("フォルダ", 23), ("状態", 31), ("未解決項目", 17) })
            fileList.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = title, FillWeight = weight, MinimumWidth = title == "状態" ? 190 : 65,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });
        fileSort = new GridSort(fileList, RenderFiles);
        fileList.CellValueNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is not { } book) return;
            e.Value = e.ColumnIndex switch
            {
                0 => Path.GetFileName(book.RelativePath),
                1 => Path.GetDirectoryName(book.RelativePath),
                2 => StatusText(BookStatus(book)),
                _ => book.Plan is null ? "—" : Pending(book).ToString()
            };
        };
        fileList.CellFormatting += (_, e) =>
        {
            if (FileAt(e.RowIndex) is not { } book || e.CellStyle is not { } style) return;
            if (BookStatus(book) is FileState.Conflict or FileState.ManualReview)
            {
                style.BackColor = Appearance.MergeConflict;
                style.SelectionBackColor = Appearance.MergeConflictSelected;
            }
        };
        fileList.CellToolTipTextNeeded += (_, e) =>
        {
            if (FileAt(e.RowIndex) is { } book) e.ToolTipText = e.ColumnIndex <= 1 ? book.RelativePath : book.Notice;
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
            (!previous.UseBase || !string.IsNullOrWhiteSpace(previous.Base)) && !string.IsNullOrWhiteSpace(previous.Production) &&
            !string.IsNullOrWhiteSpace(previous.Incoming))
        {
            basePath.Text = previous.Base ?? "";
            productionPath.Text = previous.Production;
            incomingPath.Text = previous.Incoming;
            folders.Checked = previous.Folders;
            useBase.Checked = previous.UseBase;
        }
        foreach (var box in new[] { basePath, productionPath, incomingPath })
            box.TextChanged += (_, _) => ResetInputs();
        folders.CheckedChanged += (_, _) => ResetInputs();
        useBase.CheckedChanged += (_, _) => ResetInputs();
        ResetInputs();
    }

    internal bool Busy => busy;
    internal bool UsesBase => useBase.Checked;
    internal FolderComparisonProgress? ComparisonProgress { get; private set; }
    internal string ActivityStatus => summary.Text;

    internal IReadOnlyDictionary<MergeCellAddress, MergeChoice> ChoicesFor(FolderMergeBook book)
    {
        if (!choices.TryGetValue(book.RelativePath, out var selected)) choices.Add(book.RelativePath, selected = []);
        return selected;
    }

    internal void Choose(FolderMergeBook book, MergeCellAddress cell, MergeChoice choice)
    {
        if (busy || !choices.TryGetValue(book.RelativePath, out var selected) ||
            choice is not (MergeChoice.Production or MergeChoice.Incoming) ||
            book.Plan?.Changes.FirstOrDefault(change => change.Cell == cell)?.CanChoose != true) return;
        if (conflicts[book.RelativePath].Contains(cell) && !selected.ContainsKey(cell)) unresolved--;
        selected[cell] = choice;
        RenderFiles();
        UpdateSummary();
    }

    private FolderMergeBook? Current => FileAt(fileList.CurrentCell?.RowIndex ?? -1);
    private FolderMergeBook? FileAt(int index) => index >= 0 && index < visible.Length ? visible[index] : null;
    private int Pending(FolderMergeBook book)
    {
        var selected = choices.GetValueOrDefault(book.RelativePath);
        return conflicts[book.RelativePath].Count(cell => selected?.ContainsKey(cell) != true);
    }

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
        var compareWithBase = UsesBase;
        var token = lifetime.Token;
        ResetPlan();
        FolderComparisonProgress? progress = null;
        ComparisonProgress = null;
        SetBusy(true, compareFolders ? "フォルダを比較しています…" : "ブックを比較しています…");
        using var timer = new System.Windows.Forms.Timer { Interval = 100 };
        void Report(FolderComparisonProgress value) => Volatile.Write(ref progress, value);
        void RefreshProgress()
        {
            ComparisonProgress = Volatile.Read(ref progress);
            if (ComparisonProgress is { } count)
                summary.Text = $"解析中　{count.Completed:N0} / {count.Total:N0} ファイル";
            ProgressChanged?.Invoke();
        }
        timer.Tick += (_, _) => RefreshProgress();
        timer.Start();
        try
        {
            var original = compareWithBase ? Path.GetFullPath(ReadPath(basePath)) : null;
            var production = Path.GetFullPath(ReadPath(productionPath));
            var incoming = Path.GetFullPath(ReadPath(incomingPath));
            if (compareFolders)
            {
                var result = await Task.Run(() => original is null
                    ? FolderMerger.AnalyzePair(production, incoming, token, Report)
                    : FolderMerger.Analyze(original, production, incoming, token, Report), token);
                if (IsDisposed || token.IsCancellationRequested) return;
                RefreshProgress();
                folderPlan = result;
                LoadBooks(result.Books);
            }
            else
            {
                Report(new(0, 1));
                var result = await Task.Run(() =>
                {
                    var plan = original is null ? WorkbookMerger.AnalyzePair(production, incoming, token) :
                        WorkbookMerger.Analyze(original, production, incoming, token);
                    Report(new(1, 1));
                    return plan;
                }, token);
                if (IsDisposed || token.IsCancellationRequested) return;
                RefreshProgress();
                LoadBooks([new(Path.GetFileName(production), original, production, incoming, result, null)]);
            }
            LocalSettings.Save(recentTargetsPath, new MergeTargets(original ?? ReadPath(basePath), production, incoming,
                compareFolders, compareWithBase));
            ShowInputs(false);
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error)) { if (!IsDisposed) summary.Text = error.Message; }
        finally
        {
            timer.Stop();
            ComparisonProgress = null;
            if (!IsDisposed) SetBusy(false);
        }
    }

    private sealed record MergeTargets(string? Base, string Production, string Incoming, bool Folders, bool UseBase = true);

    private void LoadBooks(IReadOnlyList<FolderMergeBook> result)
    {
        foreach (var book in result)
        {
            var cells = book.Plan?.Changes.Where(change => change.Disposition == MergeDisposition.Conflict)
                .Select(change => change.Cell).ToHashSet() ?? [];
            conflicts.Add(book.RelativePath, cells);
            unresolved += cells.Count;
            if (book.Plan is not { RequiresManualResolution: false })
                review.Add(book.RelativePath);
        }
        books = result.OrderByDescending(book => conflicts[book.RelativePath].Count > 0)
            .ThenByDescending(book => review.Contains(book.RelativePath)).ToArray();
        RenderFiles();
        UpdateSummary();
    }

    private void RenderFiles()
    {
        var old = Current;
        var filtered = books.Where(book => filter.SelectedIndex switch
        {
            1 => Pending(book) > 0 || review.Contains(book.RelativePath),
            2 => review.Contains(book.RelativePath),
            _ => true
        });
        var ordered = fileSort.Column switch
        {
            0 => fileSort.Order(filtered, book => Path.GetFileName(book.RelativePath), StringComparer.CurrentCultureIgnoreCase),
            1 => fileSort.Order(filtered, book => Path.GetDirectoryName(book.RelativePath), StringComparer.CurrentCultureIgnoreCase),
            2 => fileSort.Order(filtered, BookStatus),
            3 => fileSort.Order(filtered, book => book.Plan is null ? (int?)null : Pending(book)),
            _ => filtered
        };
        visible = ordered.ToArray();
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
        summary.Text = $"{books.Count} ファイル　未解決のコンフリクト {unresolved} 件　手動解消 {review.Count} ファイル";

    private enum FileState { Conflict, ManualReview, Resolved, Different, Same }

    private FileState BookStatus(FolderMergeBook book)
    {
        if (Pending(book) > 0) return FileState.Conflict;
        if (review.Contains(book.RelativePath)) return FileState.ManualReview;
        if (conflicts[book.RelativePath].Count > 0) return FileState.Resolved;
        return book.Plan!.PartChanges.Count + book.Plan.PairParts.Count > 0 ? FileState.Different : FileState.Same;
    }

    private static string StatusText(FileState state) => state switch
    {
        FileState.Conflict => "コンフリクトあり",
        FileState.ManualReview => "コンフリクト（手動解消）",
        FileState.Resolved => "コンフリクト解消済み",
        FileState.Different => "差分あり（コンフリクトなし）",
        _ => "同一"
    };

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
        foreach (Control control in paths.Controls)
            if (paths.GetRow(control) == 1) control.Visible = UsesBase;
        save.Text = folders.Checked ? "結果フォルダに保存…" : "別名で保存…";
        var versions = UsesBase ? "共通ベース・取り込み1・取り込み2" : "取り込み1・取り込み2";
        summary.Text = $"{versions}の{(folders.Checked ? "フォルダ" : "Excel ファイル")}を指定します。";
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
        analyze.Enabled = !busy && (!UsesBase || ReadPath(basePath).Length > 0) &&
            ReadPath(productionPath).Length > 0 && ReadPath(incomingPath).Length > 0;
        changeInputs.Visible = folderPlan is not null || books.Count > 0;
        changeInputs.Enabled = !busy;
        filter.Enabled = !busy && books.Count > 0;
        open.Enabled = preview.Enabled = !busy && Current is not null;
        save.Enabled = !busy && unresolved == 0 && review.Count == 0 && books.Count > 0;
        notice.Text = Current is { } book ? book.Notice ??
            (review.Contains(book.RelativePath) ? "「詳細を開く」で項目を確認し、Excel で修正してから再比較してください。" : "") : "";
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
