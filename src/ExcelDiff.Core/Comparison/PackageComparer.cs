using System.Buffers;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Core.Comparison;

public enum PackagePartChange { Unchanged, Added, Removed, Modified }
public enum PackagePartMergeStatus { ProductionOnly, IncomingOnly, SameChange, BothChanged }

public sealed record PackagePart(string Path, long Length, string Sha256, bool IsXml);

public sealed record PackagePartDifference(string Path, PackagePart? Before, PackagePart? After)
{
    public PackagePartChange Kind => PackageComparer.Change(Before, After);
}

public sealed record PackagePartMergeDifference(string Path, PackagePart? Base, PackagePart? Production, PackagePart? Incoming)
{
    public PackagePartChange ProductionChange => PackageComparer.Change(Base, Production);
    public PackagePartChange IncomingChange => PackageComparer.Change(Base, Incoming);
    public PackagePartMergeStatus Status =>
        ProductionChange == PackagePartChange.Unchanged ? PackagePartMergeStatus.IncomingOnly :
        IncomingChange == PackagePartChange.Unchanged ? PackagePartMergeStatus.ProductionOnly :
        PackageComparer.Change(Production, Incoming) == PackagePartChange.Unchanged ? PackagePartMergeStatus.SameChange :
        PackagePartMergeStatus.BothChanged;
}

public static class PackageComparer
{
    public static IReadOnlyList<PackagePartDifference> Compare(string before, string after, CancellationToken token = default)
        => Compare(Read(before, token), Read(after, token), token);

    internal static IReadOnlyList<PackagePartDifference> Compare(Dictionary<string, PackagePart> left,
        Dictionary<string, PackagePart> right, CancellationToken token)
    {
        var changes = new List<PackagePartDifference>();
        foreach (var path in Paths(left, right))
        {
            token.ThrowIfCancellationRequested();
            var old = left.GetValueOrDefault(path);
            var updated = right.GetValueOrDefault(path);
            if (Change(old, updated) != PackagePartChange.Unchanged) changes.Add(new(path, old, updated));
        }
        return changes.AsReadOnly();
    }

    internal static IReadOnlyList<PackagePartMergeDifference> Compare(
        Dictionary<string, PackagePart> original, Dictionary<string, PackagePart> production,
        Dictionary<string, PackagePart> incoming, CancellationToken token)
    {
        var changes = new List<PackagePartMergeDifference>();
        foreach (var path in Paths(original, production, incoming))
        {
            token.ThrowIfCancellationRequested();
            var change = new PackagePartMergeDifference(path, original.GetValueOrDefault(path),
                production.GetValueOrDefault(path), incoming.GetValueOrDefault(path));
            if (change.ProductionChange != PackagePartChange.Unchanged || change.IncomingChange != PackagePartChange.Unchanged)
                changes.Add(change);
        }
        return changes.AsReadOnly();
    }

    public static XDocument? ReadXml(string path, PackagePart? part, CancellationToken token = default)
    {
        if (part is null) return null;
        if (!part.IsXml) throw new ArgumentException("XML 部品を選択してください。", nameof(part));
        using var file = Open(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);
        var entry = zip.GetEntry(part.Path);
        if (entry is null || entry.Length != part.Length || Hash(entry, token) != part.Sha256)
            throw new InvalidDataException("解析後にファイルが変更されました。もう一度比較してください。");
        token.ThrowIfCancellationRequested();
        return XlsxReader.Part(zip, part.Path);
    }

    internal static Dictionary<string, PackagePart> Read(Stream stream, CancellationToken token)
    {
        var parts = new Dictionary<string, PackagePart>(StringComparer.Ordinal);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var types = zip.GetEntry("[Content_Types].xml") is null ? null : XlsxReader.Part(zip, "[Content_Types].xml").Root;
        var defaults = TypeMap(types, "Default", "Extension", StringComparer.OrdinalIgnoreCase);
        var overrides = TypeMap(types, "Override", "PartName", StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            token.ThrowIfCancellationRequested();
            var path = entry.FullName;
            if (path.EndsWith('/') && entry.Length == 0) continue;
            var extension = Path.GetExtension(path);
            var mime = overrides.GetValueOrDefault("/" + path) ??
                defaults.GetValueOrDefault(extension.TrimStart('.'));
            var xml = extension.Equals(".xml", StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(".rels", StringComparison.OrdinalIgnoreCase) ||
                (mime is not null && (mime.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) ||
                    mime.Equals("application/xml", StringComparison.OrdinalIgnoreCase) ||
                    mime.Equals("text/xml", StringComparison.OrdinalIgnoreCase)));
            if (!parts.TryAdd(path, new(path, entry.Length, Hash(entry, token), xml)))
                throw new InvalidDataException($"ブック内に同名の部品が複数あります。{path}");
        }
        return parts;
    }

    internal static PackagePartChange Change(PackagePart? before, PackagePart? after) =>
        before is null ? after is null ? PackagePartChange.Unchanged : PackagePartChange.Added :
        after is null ? PackagePartChange.Removed :
        before.Length == after.Length && before.Sha256 == after.Sha256 ? PackagePartChange.Unchanged : PackagePartChange.Modified;

    private static Dictionary<string, string> TypeMap(XElement? types, string element, string attribute, StringComparer comparer)
    {
        var map = new Dictionary<string, string>(comparer);
        foreach (var entry in types?.Elements(types.Name.Namespace + element) ?? [])
            if ((string?)entry.Attribute(attribute) is string key && (string?)entry.Attribute("ContentType") is string mime)
                map.TryAdd(key, mime);
        return map;
    }

    private static Dictionary<string, PackagePart> Read(string path, CancellationToken token)
    {
        using var stream = Open(path);
        return Read(stream, token);
    }

    private static FileStream Open(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read,
        bufferSize: 64 * 1024, FileOptions.SequentialScan);

    private static string Hash(ZipArchiveEntry entry, CancellationToken token)
    {
        using var stream = entry.Open();
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = ArrayPool<byte>.Shared.Rent(64 * 1024);
        try
        {
            int read;
            while ((read = stream.Read(buffer)) > 0)
            {
                token.ThrowIfCancellationRequested();
                hash.AppendData(buffer.AsSpan(0, read));
            }
            return Convert.ToHexString(hash.GetHashAndReset());
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
    }

    private static IEnumerable<string> Paths(params Dictionary<string, PackagePart>[] snapshots) =>
        snapshots.SelectMany(snapshot => snapshot.Keys).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal);
}
