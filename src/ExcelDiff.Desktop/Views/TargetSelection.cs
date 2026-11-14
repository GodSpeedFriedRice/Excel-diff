namespace ExcelDiff.Desktop.Views;

internal sealed record Targets(string Left, string Right, bool Folders, bool Recursive);

internal sealed class TargetSelection : UserControl
{
    private readonly TextBox left = new() { Dock = DockStyle.Fill, Margin = new Padding(0, 5, 8, 3), AccessibleName = "変更前のパス" };
    private readonly TextBox right = new() { Dock = DockStyle.Fill, Margin = new Padding(0, 5, 8, 3), AccessibleName = "変更後のパス" };
    private readonly RadioButton folders = new() { Text = "フォルダ比較", Checked = true, Appearance = System.Windows.Forms.Appearance.Button, AutoSize = true };
    private readonly RadioButton files = new() { Text = "ファイル比較", Appearance = System.Windows.Forms.Appearance.Button, AutoSize = true };
    private readonly CheckBox recursive = new() { Text = "サブフォルダも比較する", Checked = true, AutoSize = true };
    private readonly Label error = Appearance.Label("");
    private readonly Button compare;
    private readonly Func<Targets, Task> compareTargets;

    internal TargetSelection(Targets? previous, Func<Targets, Task> compareTargets)
    {
        this.compareTargets = compareTargets;
        Dock = DockStyle.Fill;
        BackColor = Color.White;
        var layout = Appearance.Stack();
        layout.Padding = new Padding(27, 21, 27, 14);
        var title = Appearance.Label("比較対象を選択");
        title.Font = new Font(Font.FontFamily, 16, FontStyle.Bold);
        title.Disposed += (_, _) => title.Font.Dispose();
        Appearance.Row(layout, title);
        Appearance.Row(layout, Appearance.Label("左右にフォルダ2つ、またはExcelファイル2つを指定します。"));

        var modes = Appearance.Toolbar();
        modes.BackColor = Color.White;
        modes.Controls.AddRange([folders, files]);
        Appearance.Row(layout, modes);
        AddPath(layout, "左 ・ 変更前", left);
        AddPath(layout, "右 ・ 変更後", right);
        recursive.Margin = new Padding(6, 8, 0, 0);
        Appearance.Row(layout, recursive);
        folders.CheckedChanged += (_, _) =>
        {
            recursive.Visible = folders.Checked;
            ShowError("");
        };
        left.TextChanged += (_, _) => ShowError("");
        right.TextChanged += (_, _) => ShowError("");
        error.ForeColor = Color.Firebrick;
        Appearance.Row(layout, error);
        Appearance.Row(layout, new Panel { Dock = DockStyle.Fill }, fill: true);

        var buttons = Appearance.Toolbar(FlowDirection.RightToLeft);
        buttons.BackColor = Color.White;
        compare = Appearance.Button("比較", async () => await Confirm(), primary: true);
        buttons.Controls.Add(compare);
        Appearance.Row(layout, buttons);
        Controls.Add(layout);

        if (previous is null) return;
        left.Text = previous.Left;
        right.Text = previous.Right;
        files.Checked = !previous.Folders;
        recursive.Checked = previous.Recursive;
    }

    internal void FocusPath() => left.Focus();

    internal void ShowError(string message) => error.Text = message;

    protected override bool ProcessCmdKey(ref Message message, Keys keyData)
    {
        if (keyData == Keys.Enter && (left.Focused || right.Focused))
        {
            compare.PerformClick();
            return true;
        }
        return base.ProcessCmdKey(ref message, keyData);
    }

    private void AddPath(TableLayoutPanel layout, string name, TextBox input)
    {
        Appearance.Row(layout, Appearance.Label(name));
        var row = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        row.Controls.Add(input, 0, 0);
        row.Controls.Add(Appearance.Button("参照…", () => Browse(input)), 1, 0);
        Appearance.Row(layout, row);
    }

    private void Browse(TextBox input)
    {
        if (folders.Checked)
        {
            using var picker = new FolderBrowserDialog { SelectedPath = input.Text };
            if (picker.ShowDialog(this) == DialogResult.OK) input.Text = picker.SelectedPath;
        }
        else
        {
            using var picker = new OpenFileDialog { Filter = "Excel ブック (*.xlsx)|*.xlsx", CheckFileExists = true };
            if (File.Exists(input.Text)) picker.FileName = input.Text;
            if (picker.ShowDialog(this) == DialogResult.OK) input.Text = picker.FileName;
        }
    }

    private async Task Confirm()
    {
        var a = left.Text.Trim().Trim('"');
        var b = right.Text.Trim().Trim('"');
        bool Valid(string path) => folders.Checked ? Directory.Exists(path) :
            File.Exists(path) && Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase);

        if (!Valid(a) || !Valid(b))
        {
            error.Text = folders.Checked ? "左右に存在するフォルダを指定してください。" :
                "左右に存在する .xlsx ファイルを指定してください。";
            (Valid(a) ? right : left).Focus();
            return;
        }

        error.Text = "";
        Enabled = false;
        try
        {
            await compareTargets(new Targets(Path.GetFullPath(a), Path.GetFullPath(b), folders.Checked, recursive.Checked));
        }
        finally
        {
            if (!IsDisposed) Enabled = true;
        }
    }
}
