namespace ExcelDiff.Desktop.Views;

internal sealed class ZoomBar : FlowLayoutPanel
{
    private readonly WheelZoom zoom;
    private readonly TrackBar slider = new()
    {
        Minimum = WheelZoom.Minimum, Maximum = WheelZoom.Maximum, TickStyle = TickStyle.None,
        AutoSize = false, Width = 128, Height = 27, SmallChange = 5, LargeChange = 25,
        AccessibleName = "表示倍率（1～500%）", Margin = new Padding(4, 3, 4, 1)
    };
    private readonly Label percent = new()
    {
        Width = 48, Height = 26, TextAlign = ContentAlignment.MiddleRight, ForeColor = Appearance.Ink,
        AccessibleName = "現在の倍率", Margin = new Padding(3, 3, 6, 2)
    };
    private readonly Button minus;
    private readonly Button plus;
    private readonly Button reset;

    internal ZoomBar(WheelZoom zoom)
    {
        this.zoom = zoom;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Dock = DockStyle.Fill;
        WrapContents = false;
        FlowDirection = FlowDirection.RightToLeft;
        Padding = new Padding(3, 0, 5, 1);
        BackColor = Appearance.Panel;

        minus = MakeButton("−", "縮小", () => zoom.Adjust(-10));
        plus = MakeButton("+", "拡大", () => zoom.Adjust(10));
        reset = MakeButton("100%に戻す", "表示倍率を100%に戻す", () => zoom.SetPercent(100));
        Controls.AddRange([reset, percent, plus, slider, minus]);
        UpdateValue();
        slider.ValueChanged += (_, _) => zoom.SetPercent(slider.Value);
        zoom.Changed += UpdateValue;
    }

    private static Button MakeButton(string text, string name, Action click)
    {
        var button = Appearance.Button(text, click);
        button.AccessibleName = name;
        button.MinimumSize = new Size(29, 26);
        button.Padding = new Padding(5, 0, 5, 0);
        button.Margin = new Padding(2, 2, 2, 1);
        return button;
    }

    private void UpdateValue()
    {
        slider.Value = zoom.Percent;
        percent.Text = $"{zoom.Percent}%";
        minus.Enabled = zoom.Percent > WheelZoom.Minimum;
        plus.Enabled = zoom.Percent < WheelZoom.Maximum;
        reset.Enabled = zoom.Percent != 100;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) zoom.Changed -= UpdateValue;
        base.Dispose(disposing);
    }
}
