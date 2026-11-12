using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class CompareWindow : Form
{
    private readonly TabControl tabs = new() { Dock = DockStyle.Fill, Padding = new Point(15, 5) };
    private readonly TabPage folderTab = new("フォルダ比較");
    private readonly DataGridView fileList = Appearance.Grid();
    private readonly CheckBox showSheets = new() { Text = "シートまで差分を見る", AutoSize = true, Checked = true };
    private readonly CheckBox showSame = new() { Text = "同一ファイルを表示", AutoSize = true, Checked = true };
    private readonly Label leftPath = Appearance.Label("");
    private readonly Label rightPath = Appearance.Label("");
    private readonly ToolStripStatusLabel status = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly CancellationTokenSource closing = new();
    private Targets? targets;
    private IReadOnlyList<FilePair> results = [];
    private bool busy;

    internal CompareWindow()
    {
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
        file.DropDownItems.Add("タブを閉じる", null, (_, _) => CloseSelectedTab());
        file.DropDownItems.Add("終了", null, (_, _) => Close());
        var view = new ToolStripMenuItem("表示");
        view.DropDownItems.Add("フォルダ比較", null, (_, _) => { if (tabs.TabPages.Contains(folderTab)) tabs.SelectedTab = folderTab; });
        menu.Items.AddRange([file, view]);
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
        bar.Items.Add(status);
        Appearance.Row(layout, bar);
        Controls.Add(layout);

        var folder = Appearance.Stack();
        var paths = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2 };
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        paths.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        paths.Controls.Add(leftPath, 0, 0);
        paths.Controls.Add(rightPath, 1, 0);
        Appearance.Row(folder, paths);
        Appearance.Row(folder, fileList, fill: true);
        folderTab.Controls.Add(folder);
        foreach (var (name, weight) in new[] { ("ファイル名", 22), ("フォルダ", 16), ("比較結果", 19),
            ("左更新日時", 17), ("右更新日時", 17), ("拡張子", 9) })
        {
            fileList.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = name, FillWeight = weight,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });
        }
        fileList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
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
        tabs.MouseDown += (_, e) =>
        {
            if (e.Button == MouseButtons.Middle && tabs.SelectedTab != folderTab)
                for (var i = 0; i < tabs.TabCount; i++)
                    if (tabs.GetTabRect(i).Contains(e.Location) && tabs.TabPages[i] != folderTab)
                    {
                        tabs.TabPages[i].Dispose();
                        break;
                    }
        };
        Shown += async (_, _) =>
        {
            await Choose();
            if (targets is null && !IsDisposed) Close();
        };
        FormClosed += (_, _) => closing.Cancel();
    }

    private async Task Choose()
    {
        if (busy) return;
        using var dialog = new TargetDialog(targets);
        if (dialog.ShowDialog(this) != DialogResult.OK || dialog.Selection is not Targets next) return;
        if (next.Folders) await LoadFolder(next, newSelection: true);
        else await LoadBook(next.Left, next.Right, next);
    }

    private async Task Recompare()
    {
        if (busy || targets is not Targets current) return;
        if (tabs.SelectedTab?.Tag is WorkbookView view)
        {
            await LoadBook(view.LeftPath, view.RightPath, null, tabs.SelectedTab);
            return;
        }
        if (current.Folders) await LoadFolder(current, newSelection: false);
    }

    private async Task LoadFolder(Targets current, bool newSelection)
    {
        await Run(() => FolderComparer.Compare(current.Left, current.Right, current.Recursive, closing.Token), rows =>
        {
            if (newSelection) ClearTabs();
            targets = current;
            results = rows;
            if (!tabs.TabPages.Contains(folderTab)) tabs.TabPages.Add(folderTab);
            leftPath.Text = $"左  {current.Left}";
            rightPath.Text = $"右  {current.Right}";
            tabs.SelectedTab = folderTab;
            RenderFiles();
        });
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
            }
            replace?.Dispose();
            var view = new WorkbookView(left, right, result.a, result.b, result.Diff) { Dock = DockStyle.Fill };
            if (sheet is not null) view.SelectSheet(sheet);
            var title = Path.GetFileName(left);
            if (Path.GetFileName(right) != title) title += " ↔ " + Path.GetFileName(right);
            var page = new TabPage(title) { Tag = view };
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

    private async Task Run<T>(Func<T> operation, Action<T> completed)
    {
        if (busy) return;
        busy = true;
        UseWaitCursor = true;
        showSheets.Enabled = false;
        status.Text = "比較しています…";
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
            busy = false;
            if (!IsDisposed)
            {
                UseWaitCursor = false;
                UpdateStatus();
            }
        }
    }

    private sealed record FolderRow(FilePair Pair, SheetSummary? Sheet = null);

    private void RenderFiles()
    {
        var selected = fileList.CurrentRow?.Tag as FolderRow;
        DataGridViewCell? selection = null;
        fileList.Rows.Clear();
        string? lastFolder = null;
        foreach (var pair in results)
        {
            if (pair.Status == FileStatus.Same && !showSame.Checked) continue;
            var folder = Path.GetDirectoryName(pair.RelativePath);
            if (!string.IsNullOrEmpty(folder) && folder != lastFolder)
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

    private void CloseSelectedTab()
    {
        if (tabs.SelectedTab is { Tag: WorkbookView } page) page.Dispose();
    }

    private void ClearTabs()
    {
        tabs.TabPages.Remove(folderTab);
        foreach (var page in tabs.TabPages.Cast<TabPage>().ToArray()) page.Dispose();
    }

    private void UpdateStatus()
    {
        var folder = tabs.SelectedTab == folderTab;
        showSame.Enabled = folder;
        showSheets.Enabled = folder && !busy;
        if (busy) return;
        Text = folder ? "フォルダ比較 — Excel Diff" : "ファイル比較 — Excel Diff";
        var counts = results.ToLookup(r => r.Status);
        status.Text = folder
            ? $"{results.Count} 件  ・  差分あり {counts[FileStatus.Different].Count()}  ・  同一 {counts[FileStatus.Same].Count()}" +
              $"  ・  左のみ {counts[FileStatus.LeftOnly].Count()}  ・  右のみ {counts[FileStatus.RightOnly].Count()}  ・  比較不可 {counts[FileStatus.Unreadable].Count()}"
            : (tabs.SelectedTab?.Tag as WorkbookView)?.Status ?? "";
    }
}
