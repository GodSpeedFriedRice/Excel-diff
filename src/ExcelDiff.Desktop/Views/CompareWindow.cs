using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class CompareWindow : Form
{
    private readonly ComparisonTabs tabs = new();
    private readonly TabPage folderTab = new("フォルダ比較");
    private readonly DataGridView fileList = Appearance.Grid();
    private readonly CheckBox showSheets = new() { Text = "シートまで差分を見る", AutoSize = true, Checked = true };
    private readonly CheckBox showSame = new() { Text = "同一ファイルを表示", AutoSize = true, Checked = true };
    private readonly Label leftPath = Appearance.Label("");
    private readonly Label rightPath = Appearance.Label("");
    private readonly TextBox searchText = new() { Width = 240, PlaceholderText = "ファイル名・フォルダ名", AccessibleName = "ファイル名・フォルダ名を検索" };
    private readonly FlowLayoutPanel searchBar = Appearance.Toolbar();
    private readonly ToolStripMenuItem find = new("ファイル・フォルダを検索") { ShortcutKeys = Keys.Control | Keys.F, Enabled = false };
    private readonly ToolStripStatusLabel status = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ToolStripProgressBar progressBar = new()
    {
        AutoSize = false, Width = 150, Visible = false, AccessibleName = "比較の進捗"
    };
    private readonly CancellationTokenSource closing = new();
    private Targets? targets;
    private IReadOnlyList<FilePair> results = [];
    private PreviewOptions previewOptions;
    private int sortColumn = -1;
    private SortOrder sortOrder = SortOrder.None;
    private int visibleFileCount;
    private bool busy;

    internal CompareWindow() : this(LocalSettings.PathFor("preview-settings.json")) { }

    internal CompareWindow(string previewSettingsPath)
    {
        previewOptions = LocalSettings.Load<PreviewOptions>(previewSettingsPath) ?? new();
        Text = "Excel Diff";
        ClientSize = new Size(1119, 629);
        MinimumSize = new Size(760, 470);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.White;
        ForeColor = Appearance.Ink;

        var layout = Appearance.Stack();
        var menu = new MenuStrip { Dock = DockStyle.Fill, BackColor = Color.White };
        var file = new ToolStripMenuItem("ファイル");
        file.DropDownItems.Add(new ToolStripMenuItem("開く…", null, async (_, _) => await Choose()) { ShortcutKeys = Keys.Control | Keys.O });
        file.DropDownItems.Add(new ToolStripMenuItem("タブを閉じる", null, (_, _) => CloseDetail(tabs.SelectedTab))
        {
            ShortcutKeys = Keys.Control | Keys.W
        });
        file.DropDownItems.Add("終了", null, (_, _) => Close());
        var view = new ToolStripMenuItem("表示");
        view.DropDownItems.Add("フォルダ比較", null, (_, _) => { if (tabs.TabPages.Contains(folderTab)) tabs.SelectedTab = folderTab; });
        find.Click += (_, _) => OpenSearch();
        view.DropDownItems.Add(find);
        var settings = new ToolStripMenuItem("設定");
        AddPreviewOption("周辺を表示", previewOptions.ShowNearby, (current, enabled) => current with { ShowNearby = enabled });
        AddPreviewOption("保存倍率を使う", previewOptions.UseSavedZoom, (current, enabled) => current with { UseSavedZoom = enabled });
        AddPreviewOption("グリッド線を表示", previewOptions.ShowGridlines, (current, enabled) => current with { ShowGridlines = enabled });
        menu.Items.AddRange([file, view, settings]);
        MainMenuStrip = menu;
        Appearance.Row(layout, menu);

        var toolbar = Appearance.Toolbar();
        toolbar.Controls.Add(Appearance.Button("開く…", async () => await Choose()));
        toolbar.Controls.Add(Appearance.Button("再比較", async () => await Recompare()));
        toolbar.Controls.Add(Appearance.Button("↑ 前の差分", () => MoveFile(-1)));
        toolbar.Controls.Add(Appearance.Button("↓ 次の差分", () => MoveFile(1)));
        var options = Appearance.Toolbar(FlowDirection.RightToLeft);
        foreach (var box in new[] { showSame, showSheets })
        {
            box.Margin = new Padding(12, 9, 5, 0);
            options.Controls.Add(box);
        }
        var commandRow = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, BackColor = Appearance.Panel };
        commandRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        commandRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        commandRow.Controls.Add(toolbar, 0, 0);
        commandRow.Controls.Add(options, 1, 0);
        Appearance.Row(layout, commandRow);
        Appearance.Row(layout, tabs, fill: true);
        var bar = new StatusStrip { BackColor = Appearance.Panel, Dock = DockStyle.Fill, SizingGrip = false };
        bar.Items.AddRange([progressBar, status]);
        Appearance.Row(layout, bar);
        Controls.Add(layout);

        var folder = Appearance.Stack();
        var paths = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2 };
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        paths.Controls.Add(leftPath, 0, 0);
        paths.Controls.Add(rightPath, 1, 0);
        Appearance.Row(folder, paths);
        searchBar.Visible = false;
        searchText.Margin = new Padding(6, 8, 4, 4);
        searchBar.Controls.AddRange([searchText, Appearance.Button("検索を閉じる", CloseSearch)]);
        var searchTimer = new System.Windows.Forms.Timer { Interval = 150 };
        searchTimer.Tick += (_, _) =>
        {
            searchTimer.Stop();
            RenderFiles();
        };
        searchText.TextChanged += (_, _) =>
        {
            searchTimer.Stop();
            if (searchBar.Visible) searchTimer.Start();
        };
        Disposed += (_, _) => searchTimer.Dispose();
        searchText.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Escape) return;
            e.Handled = e.SuppressKeyPress = true;
            CloseSearch();
        };
        Appearance.Row(folder, searchBar);
        Appearance.Row(folder, fileList, fill: true);
        folderTab.Controls.Add(folder);
        foreach (var (name, weight) in new[] { ("ファイル名", 22), ("フォルダ", 16), ("比較結果", 19),
            ("左更新日時", 17), ("右更新日時", 17), ("拡張子", 9) })
        {
            fileList.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = name, FillWeight = weight,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.Programmatic });
        }
        fileList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        fileList.ColumnHeaderMouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left && e.ColumnIndex >= 0) SortFiles(e.ColumnIndex);
        };
        fileList.CellDoubleClick += async (_, e) => { if (e.RowIndex >= 0) await OpenSelected(); };
        fileList.KeyDown += async (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.Handled = e.SuppressKeyPress = true;
            await OpenSelected();
        };
        showSame.CheckedChanged += (_, _) => RenderFiles();
        showSheets.CheckedChanged += (_, _) => RenderFiles();
        tabs.SelectedIndexChanged += (_, _) => UpdateStatus();
        tabs.CloseRequested += CloseDetail;
        Shown += async (_, _) =>
        {
            await Choose();
            if (targets is null && !IsDisposed) Close();
        };
        FormClosed += (_, _) => closing.Cancel();

        void AddPreviewOption(string caption, bool value, Func<PreviewOptions, bool, PreviewOptions> change)
        {
            var item = new ToolStripMenuItem(caption) { Checked = value, CheckOnClick = true };
            item.CheckedChanged += (_, _) =>
            {
                previewOptions = change(previewOptions, item.Checked);
                foreach (TabPage page in tabs.TabPages)
                    if (page.Tag is WorkbookView workbook) workbook.SetPreviewOptions(previewOptions);
                LocalSettings.Save(previewSettingsPath, previewOptions);
            };
            settings.DropDownItems.Add(item);
        }
    }

    private void OpenSearch()
    {
        if (!tabs.TabPages.Contains(folderTab)) return;
        tabs.SelectedTab = folderTab;
        searchBar.Visible = true;
        searchText.Focus();
        searchText.SelectAll();
    }

    private void CloseSearch()
    {
        searchBar.Visible = false;
        searchText.Clear();
        RenderFiles();
        fileList.Focus();
    }

    private async Task Choose()
    {
        if (busy) return;
        using var dialog = new TargetDialog(targets ?? RecentTargets.Load());
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Selection is not Targets next) return;
        if (next.Folders) await LoadFolder(next, newSelection: true);
        else await LoadBook(next.Left, next.Right, next);
    }

    private async Task Recompare()
    {
        if (busy || targets is not Targets current) return;
        if (tabs.SelectedTab?.Tag is WorkbookView view)
        {
            await LoadBook(view.LeftPath, view.RightPath, null, tabs.SelectedTab, view.SelectedSheet);
            return;
        }
        if (current.Folders) await LoadFolder(current, newSelection: false);
    }

    private async Task LoadFolder(Targets current, bool newSelection)
    {
        FolderComparisonProgress? progress = null;
        await Run(() => FolderComparer.Compare(current.Left, current.Right, current.Recursive, closing.Token,
            update => Volatile.Write(ref progress, update)), rows =>
        {
            if (newSelection)
            {
                ClearTabs();
                RecentTargets.Save(current);
                searchBar.Visible = false;
                searchText.Clear();
            }
            targets = current;
            results = rows;
            if (!tabs.TabPages.Contains(folderTab)) tabs.TabPages.Add(folderTab);
            leftPath.Text = $"左  {current.Left}";
            rightPath.Text = $"右  {current.Right}";
            tabs.SelectedTab = folderTab;
            RenderFiles();
        }, () => Volatile.Read(ref progress));
    }

    private async Task LoadBook(string left, string right, Targets? newSelection, TabPage? replace = null, string? sheet = null)
    {
        await Run(() =>
        {
            var a = XlsxReader.Read(left, closing.Token);
            var b = XlsxReader.Read(right, closing.Token);
            return (a, b, Diff: WorkbookComparer.Compare(a, b, token: closing.Token));
        }, result =>
        {
            if (newSelection is not null)
            {
                ClearTabs();
                targets = newSelection;
                RecentTargets.Save(newSelection);
            }
            if (replace?.IsDisposed == true) return;
            replace?.Dispose();
            var view = new WorkbookView(left, right, result.a, result.b, result.Diff) { Dock = DockStyle.Fill };
            view.SetPreviewOptions(previewOptions);
            if (sheet is not null) view.SelectSheet(sheet);
            var title = Path.GetFileName(left);
            if (Path.GetFileName(right) != title) title += " ↔ " + Path.GetFileName(right);
            var page = new TabPage(title) { Tag = view, ToolTipText = title };
            page.Controls.Add(view);
            tabs.TabPages.Add(page);
            tabs.SelectedTab = page;
            view.StatusChanged += (_, _) => UpdateStatus();
        });
    }

    private async Task OpenSelected()
    {
        if (busy || fileList.CurrentRow?.Tag is not FolderRow selected) return;
        var pair = selected.Pair;
        if (pair.Error is not null)
        {
            MessageBox.Show(this, pair.Error, "比較できません", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (pair.Left is not FileInfo a || pair.Right is not FileInfo b) return;
        foreach (TabPage page in tabs.TabPages)
            if (page.Tag is WorkbookView view && view.LeftPath == a.FullName && view.RightPath == b.FullName)
            {
                tabs.SelectedTab = page;
                if (selected.Sheet is not null) view.SelectSheet(selected.Sheet.Name);
                return;
            }
        await LoadBook(a.FullName, b.FullName, null, sheet: selected.Sheet?.Name);
    }

    private async Task Run<T>(Func<T> operation, Action<T> completed, Func<FolderComparisonProgress?>? folderProgress = null)
    {
        if (busy) return;
        busy = true;
        UseWaitCursor = true;
        showSheets.Enabled = false;
        status.Text = "比較しています…";
        using var progressTimer = folderProgress is null ? null : new System.Windows.Forms.Timer { Interval = 100 };
        if (progressTimer is not null)
        {
            status.Text = "比較対象ファイルを確認しています…";
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.Visible = true;
            progressTimer.Tick += (_, _) => ShowProgress(folderProgress!());
            progressTimer.Start();
        }
        try
        {
            var result = await Task.Run(operation, closing.Token);
            if (!IsDisposed) completed(result);
        }
        catch (OperationCanceledException) when (closing.IsCancellationRequested) { }
        catch (Exception e) when (FolderComparer.IsReadError(e))
        {
            if (!IsDisposed) MessageBox.Show(this, e.Message, "比較できません", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            progressTimer?.Stop();
            busy = false;
            if (!IsDisposed)
            {
                progressBar.Visible = false;
                UseWaitCursor = false;
                UpdateStatus();
            }
        }
    }

    private void ShowProgress(FolderComparisonProgress? progress)
    {
        if (progress is null) return;
        var percent = progress.Total == 0 ? 0 : (int)(100L * progress.Completed / progress.Total);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Value = percent;
        status.Text = $"比較中  {progress.Completed:N0} / {progress.Total:N0} 件（{percent}%）";
    }

    private sealed record FolderRow(FilePair Pair, SheetSummary? Sheet = null);

    private void SortFiles(int column)
    {
        if (sortColumn >= 0) fileList.Columns[sortColumn].HeaderCell.SortGlyphDirection = SortOrder.None;
        sortOrder = sortColumn == column && sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
        sortColumn = column;
        fileList.Columns[column].HeaderCell.SortGlyphDirection = sortOrder;
        RenderFiles();
    }

    private IEnumerable<FilePair> VisibleFiles(string query)
    {
        var pairs = results.Where(pair => (showSame.Checked || pair.Status != FileStatus.Same) &&
            pair.RelativePath.Contains(query, StringComparison.OrdinalIgnoreCase));
        var comparer = StringComparer.CurrentCultureIgnoreCase;
        IOrderedEnumerable<FilePair> Order<TKey>(Func<FilePair, TKey> key, IComparer<TKey>? order = null) =>
            sortOrder == SortOrder.Descending ? pairs.OrderByDescending(key, order) : pairs.OrderBy(key, order);

        return sortColumn switch
        {
            0 => Order(pair => Path.GetFileName(pair.RelativePath), comparer),
            1 => Order(pair => Path.GetDirectoryName(pair.RelativePath), comparer),
            2 => Order(pair => StatusName(pair.Status), comparer),
            3 => Order(pair => pair.Left?.LastWriteTimeUtc),
            4 => Order(pair => pair.Right?.LastWriteTimeUtc),
            5 => Order(pair => Path.GetExtension(pair.RelativePath), comparer),
            _ => pairs
        };
    }

    private void RenderFiles()
    {
        var selected = fileList.CurrentRow?.Tag as FolderRow;
        DataGridViewCell? selection = null;
        fileList.Rows.Clear();
        visibleFileCount = 0;
        var query = searchText.Text.Trim().Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        string? lastFolder = null;
        foreach (var pair in VisibleFiles(query))
        {
            visibleFileCount++;
            var folder = Path.GetDirectoryName(pair.RelativePath);
            if (sortColumn is -1 or 1 && !string.IsNullOrEmpty(folder) && folder != lastFolder)
            {
                var header = fileList.Rows[fileList.Rows.Add(folder)];
                header.DefaultCellStyle.BackColor = Appearance.Panel;
            }
            lastFolder = folder;
            var row = fileList.Rows[fileList.Rows.Add(Path.GetFileName(pair.RelativePath), folder ?? "", StatusName(pair.Status),
                Timestamp(pair.Left), Timestamp(pair.Right), "xlsx")];
            row.Tag = new FolderRow(pair);
            row.Cells[2].ToolTipText = pair.Error;
            row.DefaultCellStyle.BackColor = pair.Status switch
            {
                FileStatus.Different => Appearance.Changed,
                FileStatus.LeftOnly or FileStatus.RightOnly or FileStatus.Unreadable => Appearance.Panel,
                _ => Color.White
            };
            var sameFile = pair.RelativePath == selected?.Pair.RelativePath;
            if (sameFile) selection = row.Cells[0];
            if (!showSheets.Checked) continue;

            foreach (var sheet in pair.Sheets)
            {
                var summary = sheet.Status == FileStatus.Different
                    ? $"差分あり ({sheet.DifferenceCount} 件)"
                    : StatusName(sheet.Status);
                var child = fileList.Rows[fileList.Rows.Add($"    ↪ {sheet.Name}", "", summary, "", "", "")];
                child.Tag = new FolderRow(pair, sheet);
                child.DefaultCellStyle.BackColor = sheet.Status switch
                {
                    FileStatus.Different => Appearance.Changed,
                    FileStatus.LeftOnly => Appearance.Removed,
                    FileStatus.RightOnly => Appearance.Added,
                    _ => Color.White
                };
                if (sameFile && sheet.Name == selected?.Sheet?.Name) selection = child.Cells[0];
            }
        }
        if (selection is not null) fileList.CurrentCell = selection;
        UpdateStatus();
    }

    private static string Timestamp(FileInfo? file) => file is null ? "—" : file.LastWriteTime.ToString("yyyy/MM/dd HH:mm");

    private static string StatusName(FileStatus state) => state switch
    {
        FileStatus.Different => "差分あり", FileStatus.Same => "同一", FileStatus.LeftOnly => "左のみに存在",
        FileStatus.RightOnly => "右のみに存在", _ => "比較できません"
    };

    private void MoveFile(int direction)
    {
        if (tabs.SelectedTab?.Tag is WorkbookView view)
        {
            view.MoveDifference(direction);
            return;
        }
        var start = fileList.CurrentRow?.Index ?? (direction > 0 ? -1 : fileList.Rows.Count);
        for (var index = start + direction; index >= 0 && index < fileList.Rows.Count; index += direction)
            if (fileList.Rows[index].Tag is FolderRow { Pair.Status: not FileStatus.Same, Sheet: null })
            {
                fileList.CurrentCell = fileList.Rows[index].Cells[0];
                return;
            }
    }

    private void CloseDetail(TabPage? page)
    {
        if (page?.Tag is not WorkbookView || !tabs.TabPages.Contains(page)) return;
        var wasSelected = tabs.SelectedTab == page;
        page.Dispose();
        if (wasSelected && tabs.TabCount == 1 && tabs.TabPages.Contains(folderTab)) tabs.SelectedTab = folderTab;
        UpdateStatus();
    }

    private void ClearTabs()
    {
        tabs.TabPages.Remove(folderTab);
        foreach (var page in tabs.TabPages.Cast<TabPage>().ToArray()) page.Dispose();
    }

    private void UpdateStatus()
    {
        var folder = tabs.SelectedTab == folderTab;
        find.Enabled = tabs.TabPages.Contains(folderTab);
        showSame.Enabled = folder;
        showSheets.Enabled = folder && !busy;
        if (busy) return;
        Text = folder ? "フォルダ比較 — Excel Diff" : "ファイル比較 — Excel Diff";
        var counts = results.ToLookup(r => r.Status);
        var searchCount = string.IsNullOrWhiteSpace(searchText.Text) ? "" : $"検索結果 {visibleFileCount} / ";
        status.Text = folder
            ? $"{searchCount}{results.Count} 件  ・  差分あり {counts[FileStatus.Different].Count()}  ・  同一 {counts[FileStatus.Same].Count()}" +
              $"  ・  左のみ {counts[FileStatus.LeftOnly].Count()}  ・  右のみ {counts[FileStatus.RightOnly].Count()}  ・  比較不可 {counts[FileStatus.Unreadable].Count()}"
            : (tabs.SelectedTab?.Tag as WorkbookView)?.Status ?? "";
    }
}
