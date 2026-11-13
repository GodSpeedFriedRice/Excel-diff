using Theme = ExcelDiff.Desktop.Views.Appearance;

namespace ExcelDiff.Desktop.Views;

internal sealed class ComparisonTabs : TabControl
{
    private const int LeftButtonDown = 0x0201;
    private const int LeftButtonUp = 0x0202;
    private const int LeftButtonDoubleClick = 0x0203;
    private TabPage? pressed;
    private TabPage? hoveredClose;

    internal ComparisonTabs()
    {
        Dock = DockStyle.Fill;
        DrawMode = TabDrawMode.OwnerDrawFixed;
        SizeMode = TabSizeMode.Fixed;
        ItemSize = new Size(170, 29);
        ShowToolTips = true;
    }

    internal event Action<TabPage>? CloseRequested;

    protected override void OnDrawItem(DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= TabCount) return;
        var page = TabPages[e.Index];
        var bounds = GetTabRect(e.Index);
        var selected = SelectedIndex == e.Index;
        using var background = new SolidBrush(selected ? Color.White : Theme.Panel);
        e.Graphics.FillRectangle(background, bounds);
        var close = CloseBounds(bounds);
        var label = Rectangle.FromLTRB(bounds.Left + 8, bounds.Top,
            page.Tag is WorkbookView ? close.Left - 4 : bounds.Right - 6, bounds.Bottom);
        TextRenderer.DrawText(e.Graphics, page.Text, Font, label, Theme.Ink,
            TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        if (page.Tag is WorkbookView)
            TextRenderer.DrawText(e.Graphics, "×", Font, close, Theme.Ink,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        SetHoveredClose(CloseAt(e.Location));
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        SetHoveredClose(null);
        base.OnMouseLeave(e);
    }

    protected override void OnControlRemoved(ControlEventArgs e)
    {
        if (e.Control == hoveredClose) SetHoveredClose(null);
        base.OnControlRemoved(e);
    }

    private void SetHoveredClose(TabPage? page)
    {
        if (page == hoveredClose) return;
        hoveredClose = page;
        Cursor = page is null ? Cursors.Default : Cursors.Hand;
    }

    protected override void WndProc(ref Message message)
    {
        if (message.Msg is LeftButtonDown or LeftButtonDoubleClick or LeftButtonUp)
        {
            var coordinates = message.LParam.ToInt64();
            var point = new Point(unchecked((short)coordinates), unchecked((short)(coordinates >> 16)));
            if (message.Msg == LeftButtonUp && pressed is TabPage page)
            {
                pressed = null;
                Capture = false;
                if (page == CloseAt(point)) CloseRequested?.Invoke(page);
                return;
            }
            if (message.Msg != LeftButtonUp && CloseAt(point) is TabPage target)
            {
                pressed = target;
                Capture = true;
                return;
            }
        }
        base.WndProc(ref message);
    }

    protected override void OnMouseCaptureChanged(EventArgs e)
    {
        if (!Capture) pressed = null;
        base.OnMouseCaptureChanged(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button != MouseButtons.Middle) return;
        for (var index = 0; index < TabCount; index++)
        {
            if (TabPages[index].Tag is not WorkbookView || !GetTabRect(index).Contains(e.Location)) continue;
            CloseRequested?.Invoke(TabPages[index]);
            return;
        }
    }

    private TabPage? CloseAt(Point point)
    {
        for (var index = 0; index < TabCount; index++)
            if (TabPages[index].Tag is WorkbookView && CloseBounds(GetTabRect(index)).Contains(point))
                return TabPages[index];
        return null;
    }

    private Rectangle CloseBounds(Rectangle tab)
    {
        var width = Math.Max(21, tab.Height - 4);
        return new Rectangle(tab.Right - width - 4, tab.Top + 2, width, tab.Height - 4);
    }
}
