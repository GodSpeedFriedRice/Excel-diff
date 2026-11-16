namespace ExcelDiff.Desktop.Views;

internal sealed class WheelZoom
{
    internal const int Minimum = 1;
    internal const int Maximum = 500;
    private int remainder;

    internal int Percent { get; private set; } = 100;
    internal double Scale => Percent / 100.0;
    internal event Action? Changed;

    internal void Apply(int delta)
    {
        remainder += delta;
        var steps = remainder / SystemInformation.MouseWheelScrollDelta;
        remainder %= SystemInformation.MouseWheelScrollDelta;
        Adjust(steps * 10);
    }

    internal void Adjust(int step) => SetPercent(Percent + step);

    internal void SetPercent(int percent)
    {
        var next = Math.Clamp(percent, Minimum, Maximum);
        if (next == Percent) return;
        Percent = next;
        Changed?.Invoke();
    }
}

internal sealed class DifferenceGrid : DataGridView
{
    private double scale = 1;
    private Font? zoomFont;

    internal DifferenceGrid()
    {
        RowHeightInfoNeeded += (_, e) => e.Height = Math.Max(3, (int)Math.Round(28 * scale));
    }

    internal event Action<int>? ZoomRequested;

    internal void SetZoom(double next)
    {
        if (scale == next) return;
        scale = next;
        var standard = Parent?.Font ?? DefaultFont;
        var previous = zoomFont;
        zoomFont = new Font(standard.FontFamily, Math.Max(1, (float)(standard.SizeInPoints * scale)), standard.Style, GraphicsUnit.Point);
        SuspendLayout();
        try
        {
            Font = zoomFont;
            ColumnHeadersDefaultCellStyle.Font = zoomFont;
            RowTemplate.Height = Math.Max(3, (int)Math.Round(28 * scale));
            ColumnHeadersHeight = Math.Max(4, (int)Math.Round(30 * scale));
            if (RowCount > 0) UpdateRowHeightInfo(0, true);
        }
        finally
        {
            ResumeLayout();
            previous?.Dispose();
        }
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) == 0)
        {
            base.OnMouseWheel(e);
            return;
        }
        if (e is HandledMouseEventArgs handled) handled.Handled = true;
        ZoomRequested?.Invoke(e.Delta);
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing) zoomFont?.Dispose();
    }
}
