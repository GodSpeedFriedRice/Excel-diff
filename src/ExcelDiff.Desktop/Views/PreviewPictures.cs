using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO.Compression;
using System.Windows.Media.Imaging;
using ExcelDiff.Core.Workbooks;
using PixelFormats = System.Windows.Media.PixelFormats;

namespace ExcelDiff.Desktop.Views;

internal sealed class PreviewPictures
{
    private const string MaxEdgeVariable = "EXCEL_DIFF_IMAGE_MAX_EDGE";
    private const int DefaultMaxEdge = 512;
    private const long CacheBudget = 32 * 1024 * 1024;
    private const int CacheEntryLimit = 128;
    private const int MaxImageBytes = 32 * 1024 * 1024;
    private const double EmuPerPixel = 9525;
    private static readonly object Gate = new();
    private static readonly int MaxEdge = ReadMaxEdge(Environment.GetEnvironmentVariable(MaxEdgeVariable));
    private static readonly Dictionary<PictureAsset, CacheEntry> Cache = [];
    private static long cacheBytes;
    private static long currentAccess;
    private readonly Placement[] drawings;
    private readonly double zoom;

    internal PreviewPictures(Sheet sheet, PreviewAxis rows, PreviewAxis columns, double zoom)
    {
        this.zoom = zoom;
        var right = columns.Position(columns.Limit + 1);
        var bottom = rows.Position(rows.Limit + 1);
        drawings = sheet.Pictures.Select(picture => Place(picture.From, picture.To, picture.Width, picture.Height,
                picture.Order, picture.Asset, null))
            .Concat(sheet.Shapes.Select(shape => Place(shape.From, shape.To, shape.Width, shape.Height, shape.Order, null, shape)))
            .Where(drawing => drawing.VisibleRight > 0 && drawing.VisibleBottom > 0 && drawing.VisibleLeft < right &&
                drawing.VisibleTop < bottom && drawing.HasArea)
            .OrderBy(drawing => drawing.Order).ToArray();
        LastRow = drawings.Length == 0 ? 0 : rows.AtPixel(drawings.Max(drawing => drawing.VisibleBottom) - 1);
        LastColumn = drawings.Length == 0 ? 0 : columns.AtPixel(drawings.Max(drawing => drawing.VisibleRight) - 1);

        Placement Place(PictureMarker from, PictureMarker? to, long width, long height, int order, PictureAsset? asset, SheetShape? shape)
        {
            var left = X(from);
            var top = Y(from);
            var endX = to is { } end ? X(end) : left + Pixels(width);
            var endY = to is { } finish ? Y(finish) : top + Pixels(height);
            if (shape is null) return new(left, top, endX, endY, left, top, endX, endY, order, asset, null);
            var w = Math.Max(0, endX - left);
            var h = Math.Max(0, endY - top);
            var rotation = shape.Rotation * Math.PI / 180;
            var cos = Math.Abs(Math.Cos(rotation));
            var sin = Math.Abs(Math.Sin(rotation));
            var margin = Math.Clamp(Pixels(shape.Outline.Width), 2, 64) * 3;
            var dx = (long)Math.Ceiling((w * cos + h * sin - w) / 2) + margin;
            var dy = (long)Math.Ceiling((w * sin + h * cos - h) / 2) + margin;
            return new(left, top, left + w, top + h, left - dx, top - dy, left + w + dx, top + h + dy, order, null, shape);
        }

        long Pixels(long emu) => (long)Math.Round(Math.Clamp(emu / EmuPerPixel * zoom, -int.MaxValue, int.MaxValue));
        long X(PictureMarker marker) => columns.Position(marker.Column) + Pixels(marker.OffsetX);
        long Y(PictureMarker marker) => rows.Position(marker.Row) + Pixels(marker.OffsetY);
    }

    internal int LastRow { get; }
    internal int LastColumn { get; }

    internal void Draw(Graphics graphics, long x, long y, int size)
    {
        foreach (var picture in drawings)
        {
            if (picture.VisibleRight <= x || picture.VisibleBottom <= y || picture.VisibleLeft >= x + size || picture.VisibleTop >= y + size)
                continue;

            var bounds = RectangleF.FromLTRB(picture.Left - x, picture.Top - y, picture.Right - x, picture.Bottom - y);
            if (picture.Shape is { } shape)
            {
                PreviewShapes.Draw(graphics, bounds, shape, zoom);
                continue;
            }
            if (picture.Asset is not { } asset || !TryDraw(graphics, bounds, asset))
            {
                using var pen = new Pen(Color.FromArgb(109, 124, 138)) { DashStyle = DashStyle.Dash };
                graphics.DrawRectangle(pen, bounds.X, bounds.Y, bounds.Width, bounds.Height);
            }
        }
    }

    internal static int ReadMaxEdge(string? value) =>
        int.TryParse(value, out var edge) && edge is >= 1 and <= 2048 ? edge : DefaultMaxEdge;

    private static bool TryDraw(Graphics graphics, RectangleF bounds, PictureAsset asset)
    {
        lock (Gate)
        {
            if (Cache.TryGetValue(asset, out var hit)) return DrawCached(hit);
            var bitmap = Decode(asset, MaxEdge);
            var entry = new CacheEntry(bitmap);
            while (Cache.Count >= CacheEntryLimit || cacheBytes + entry.Bytes > CacheBudget)
            {
                var oldest = Cache.MinBy(pair => pair.Value.LastUse);
                Cache.Remove(oldest.Key);
                cacheBytes -= oldest.Value.Bytes;
                oldest.Value.Bitmap?.Dispose();
            }
            Cache.Add(asset, entry);
            cacheBytes += entry.Bytes;
            return DrawCached(entry);
        }

        bool DrawCached(CacheEntry entry)
        {
            entry.LastUse = ++currentAccess;
            if (entry.Bitmap is not { } image) return false;
            var interpolation = graphics.InterpolationMode;
            graphics.InterpolationMode = InterpolationMode.Bilinear;
            try { graphics.DrawImage(image, bounds); }
            finally { graphics.InterpolationMode = interpolation; }
            return true;
        }
    }

    internal static Bitmap? Decode(PictureAsset asset, int edge)
    {
        var extension = Path.GetExtension(asset.PartPath).ToLowerInvariant();
        if (extension is not (".png" or ".jpg" or ".jpeg")) return null;
        try
        {
            using var file = new FileStream(asset.WorkbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            if (file.Length != asset.WorkbookLength || File.GetLastWriteTimeUtc(asset.WorkbookPath) != asset.LastWriteTimeUtc)
                return null;
            using var zip = new ZipArchive(file, ZipArchiveMode.Read);
            if (zip.GetEntry(asset.PartPath) is not { Length: > 0 and <= MaxImageBytes } part) return null;
            using var packed = part.Open();
            using var data = new MemoryStream((int)part.Length);
            var buffer = new byte[81920];
            int read;
            while ((read = packed.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (data.Length + read > MaxImageBytes) return null;
                data.Write(buffer, 0, read);
            }
            data.Position = 0;
            var decoder = BitmapDecoder.Create(data, BitmapCreateOptions.DelayCreation, BitmapCacheOption.OnDemand);
            if (decoder is not (PngBitmapDecoder or JpegBitmapDecoder)) return null;
            var frame = decoder.Frames[0];
            var w = frame.PixelWidth;
            var h = frame.PixelHeight;
            if (w <= 0 || h <= 0 || (long)w * h > 64_000_000) return null;
            var source = (BitmapSource)frame;
            if (Math.Max(w, h) > edge)
            {
                data.Position = 0;
                var reduced = new BitmapImage();
                reduced.BeginInit();
                reduced.CacheOption = BitmapCacheOption.OnLoad;
                reduced.StreamSource = data;
                if (w >= h) reduced.DecodePixelWidth = edge;
                else reduced.DecodePixelHeight = edge;
                reduced.EndInit();
                source = reduced;
            }
            var pixels = new FormatConvertedBitmap(source, PixelFormats.Pbgra32, null, 0);
            var bitmap = new Bitmap(pixels.PixelWidth, pixels.PixelHeight, PixelFormat.Format32bppPArgb);
            try
            {
                var region = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
                var bits = bitmap.LockBits(region, ImageLockMode.WriteOnly, bitmap.PixelFormat);
                try { pixels.CopyPixels(System.Windows.Int32Rect.Empty, bits.Scan0, bits.Stride * bitmap.Height, bits.Stride); }
                finally { bitmap.UnlockBits(bits); }
                return bitmap;
            }
            catch
            {
                bitmap.Dispose();
                throw;
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or NotSupportedException or
            InvalidDataException or FormatException or InvalidOperationException or ArgumentException or System.Runtime.InteropServices.ExternalException)
        {
            return null;
        }
    }

    private sealed class CacheEntry(Bitmap? bitmap)
    {
        internal Bitmap? Bitmap { get; } = bitmap;
        internal long LastUse { get; set; }
        internal long Bytes { get; } = bitmap is null ? 0 : (long)bitmap.Width * bitmap.Height * 4;
    }

    private readonly record struct Placement(long Left, long Top, long Right, long Bottom,
        long VisibleLeft, long VisibleTop, long VisibleRight, long VisibleBottom, int Order, PictureAsset? Asset, SheetShape? Shape)
    {
        internal bool HasArea => Right > Left && Bottom > Top ||
            (Right > Left || Bottom > Top) &&
            Shape?.Geometry is "line" or "lineInv" or "straightConnector1" or "bentConnector2" or "bentConnector3";
    }
}
