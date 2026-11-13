namespace ExcelDiff.Desktop.Views;

internal static class Appearance
{
    internal static readonly Color Panel = Color.FromArgb(241, 244, 248);
    internal static readonly Color Line = Color.FromArgb(220, 228, 238);
    internal static readonly Color Ink = Color.FromArgb(30, 43, 61);
    internal static readonly Color Muted = Color.FromArgb(75, 93, 116);
    internal static readonly Color Blue = Color.FromArgb(23, 77, 159);
    internal static readonly Color Changed = Color.FromArgb(255, 239, 199);
    internal static readonly Color Added = Color.FromArgb(222, 242, 232);
    internal static readonly Color Removed = Color.FromArgb(250, 230, 235);

    internal static Button Button(string text, Action action, bool primary = false)
    {
        var button = new Button
        {
            Text = text, AutoSize = true, MinimumSize = new Size(74, 31), Padding = new Padding(8, 0, 8, 0),
            FlatStyle = FlatStyle.Flat, BackColor = primary ? Blue : Color.White,
            ForeColor = primary ? Color.White : Ink, UseVisualStyleBackColor = false
        };
        button.FlatAppearance.BorderColor = primary ? Blue : Line;
        button.Click += (_, _) => action();
        return button;
    }

    internal static FlowLayoutPanel Toolbar(FlowDirection direction = FlowDirection.LeftToRight) => new()
    {
        AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill,
        WrapContents = false, Padding = new Padding(6, 3, 6, 3), FlowDirection = direction, BackColor = Panel
    };

    internal static DataGridView Grid() => ConfigureGrid(new DataGridView());

    internal static T ConfigureGrid<T>(T grid) where T : DataGridView
    {
        grid.Dock = DockStyle.Fill;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.MultiSelect = false;
        grid.RowHeadersVisible = false;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.None;
        grid.GridColor = Line;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersHeight = 30;
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        grid.RowTemplate.Height = 28;
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Panel, ForeColor = Muted };
        grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            ForeColor = Ink, SelectionBackColor = Color.FromArgb(215, 229, 251), SelectionForeColor = Ink
        };
        return grid;
    }

    internal static Label Label(string text) => new()
    {
        Text = text, AutoSize = true, ForeColor = Ink, Margin = new Padding(6, 8, 6, 6)
    };

    internal static TableLayoutPanel Stack() => new() { Dock = DockStyle.Fill, ColumnCount = 1, Margin = Padding.Empty };

    internal static void Row(TableLayoutPanel layout, Control control, bool fill = false)
    {
        layout.RowStyles.Add(new RowStyle(fill ? SizeType.Percent : SizeType.AutoSize, fill ? 100 : 0));
        layout.Controls.Add(control, 0, layout.RowCount++);
    }
}
