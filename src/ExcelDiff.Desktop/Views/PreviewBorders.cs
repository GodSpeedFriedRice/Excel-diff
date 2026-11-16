namespace ExcelDiff.Desktop.Views;

internal sealed class PreviewBorders(long originX, long originY, bool showGridlines)
{
    private readonly Dictionary<Edge, PreviewStyle.Stroke> strokes = [];
    private readonly HashSet<Edge> grid = [];
    private readonly List<(Rectangle Bounds, PreviewStyle Style)> diagonals = [];

    internal void Add(Rectangle bounds, PreviewStyle style, BorderSides sides, PreviewStyle? anchor = null)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0) return;
        var filled = (anchor ?? style).HasFill;
        AddSide(BorderSides.Left, new(true, bounds.Left, bounds.Top, bounds.Bottom));
        AddSide(BorderSides.Top, new(false, bounds.Top, bounds.Left, bounds.Right));
        AddSide(BorderSides.Right, new(true, bounds.Right, bounds.Top, bounds.Bottom));
        AddSide(BorderSides.Bottom, new(false, bounds.Bottom, bounds.Left, bounds.Right));
        if ((sides & style.Sides & BorderSides.Diagonal) != 0) diagonals.Add((bounds, style));

        void AddSide(BorderSides side, Edge edge)
        {
            if ((sides & side) == 0) return;
            if (showGridlines && !filled) grid.Add(edge);
            var stroke = style.Edge(side);
            if (stroke.Width == 0 && anchor is not null) stroke = anchor.Edge(side);
            if (stroke.Width == 0) return;
            if (!strokes.TryGetValue(edge, out var existing) || stroke.Priority > existing.Priority)
                strokes[edge] = stroke;
        }
    }

    internal void DrawGrid(Graphics graphics, Pen pen)
    {
        pen.Color = Appearance.Line;
        foreach (var edge in grid)
            if (!strokes.ContainsKey(edge)) DrawEdge(graphics, pen, edge, 0);
    }

    internal void DrawBorders(Graphics graphics, Pen pen)
    {
        foreach (var (edge, stroke) in strokes)
        {
            pen.Color = stroke.Color;
            pen.DashStyle = stroke.Dash;
            var isOuter = edge.Position + (edge.Vertical ? originX : originY) == 0;
            var start = isOuter ? 0 : -(stroke.Width - 1) / 2;
            for (var offset = 0; offset < stroke.Width; offset++)
                if (!stroke.Double || offset != 1) DrawEdge(graphics, pen, edge, start + offset);
        }
        foreach (var (bounds, style) in diagonals) style.DrawDiagonal(graphics, pen, bounds);
    }

    private void DrawEdge(Graphics graphics, Pen pen, Edge edge, int offset)
    {
        var origin = edge.Vertical ? originX : originY;
        var position = edge.Position + (edge.Position + origin == 0 ? 0 : -1) + offset;
        if (edge.Vertical) graphics.DrawLine(pen, position, edge.First, position, edge.End - 1);
        else graphics.DrawLine(pen, edge.First, position, edge.End - 1, position);
    }

    private readonly record struct Edge(bool Vertical, int Position, int First, int End);
}
