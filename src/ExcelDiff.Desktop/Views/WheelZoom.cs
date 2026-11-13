namespace ExcelDiff.Desktop.Views;

internal sealed class WheelZoom
{
    private int remainder;

    internal int Percent { get; private set; } = 100;
    internal double Scale => Percent / 100.0;

    internal bool Apply(int delta)
    {
        remainder += delta;
        var steps = remainder / SystemInformation.MouseWheelScrollDelta;
        remainder %= SystemInformation.MouseWheelScrollDelta;
        var next = Math.Clamp(Percent + steps * 10, 50, 200);
        if (next == Percent) return false;
        Percent = next;
        return true;
    }
}

internal sealed class DifferenceGrid : DataGridView
{
    internal event Action<int>? ZoomRequested;

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

    internal void RefreshRowHeight()
    {
        if (RowCount > 0) UpdateRowHeightInfo(0, true);
    }
}
