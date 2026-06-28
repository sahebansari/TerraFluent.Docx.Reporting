using System.Security.Cryptography;
using TerraFluent.Docx.Reporting.Core.Elements;

namespace TerraFluent.Docx.Reporting.Ooxml;

internal sealed record ImageEntry(string RelationshipId, string MediaPath, string ContentType);
internal sealed record HyperlinkEntry(string RelationshipId, string Url);
internal sealed record ChartEntry(string RelationshipId, string ChartPath, ChartElement Chart);
internal sealed record ChartPartEntry(string ChartPath, ChartElement Chart);

internal sealed record HeaderFooterPart(
    bool IsHeader,
    string ReferenceType,
    string RelationshipId,
    string PartName,
    string Xml,
    IReadOnlyList<ImageEntry> Images,
    IReadOnlyList<HyperlinkEntry> Hyperlinks,
    IReadOnlyList<ChartEntry> Charts);

internal sealed record SectionPart(IReadOnlyList<HeaderFooterPart> HeaderFooterParts);

internal sealed record MediaEntry(string Key, string? FilePath, byte[]? Bytes, string MediaPath, string ContentType);
internal sealed record NoteEntry(int Id, TextElement Text);

internal sealed record NumberingLevelDefinition(
    int Level,
    string NumberFormat,
    string LevelText,
    int LeftIndent,
    int HangingIndent,
    string? FontFamily);

internal sealed record AbstractNumberingDefinition(
    int AbstractNumberId,
    IReadOnlyList<NumberingLevelDefinition> Levels);

internal sealed record NumberingInstance(int NumberId, int AbstractNumberId);

internal sealed class DocumentRenderContext
{
    private int _nextMediaId = 1;
    private int _nextDocPrId = 1;
    private int _nextBookmarkId = 1;
    private int _nextFootnoteId = 1;
    private int _nextEndnoteId = 1;
    private int _nextAbstractNumberId = 1;
    private int _nextNumberId = 1;
    private int _nextChartId = 1;
    private int _nextFigureNumber = 1;
    private int _nextTableNumber = 1;

    private readonly Dictionary<string, MediaEntry> _mediaByKey = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, AbstractNumberingDefinition> _abstractNumberingByKey = new(StringComparer.Ordinal);
    private readonly List<NumberingInstance> _numberingInstances = [];
    private readonly List<ChartPartEntry> _charts = [];

    public RelationshipScope DocumentRelationships { get; }
    public RelationshipScope FootnoteRelationships { get; }
    public RelationshipScope EndnoteRelationships { get; }

    public DocumentRenderContext()
    {
        // rId1=styles, rId2=settings, rId3=numbering are reserved in word/document.xml.rels.
        DocumentRelationships = new RelationshipScope(this, firstRelationshipId: 4);
        FootnoteRelationships = new RelationshipScope(this, firstRelationshipId: 1);
        EndnoteRelationships = new RelationshipScope(this, firstRelationshipId: 1);
    }

    public IReadOnlyCollection<MediaEntry> Media => _mediaByKey.Values;
    public IReadOnlyList<ImageEntry> Images => DocumentRelationships.Images;
    public IReadOnlyList<HyperlinkEntry> Hyperlinks => DocumentRelationships.Hyperlinks;
    public IReadOnlyList<ChartEntry> Charts => DocumentRelationships.Charts;
    public IReadOnlyList<ChartPartEntry> ChartParts => _charts;
    public IReadOnlyCollection<AbstractNumberingDefinition> AbstractNumberingDefinitions => _abstractNumberingByKey.Values;
    public IReadOnlyList<NumberingInstance> NumberingInstances => _numberingInstances;
    public List<NoteEntry> Footnotes { get; } = [];
    public List<NoteEntry> Endnotes { get; } = [];
    public List<SectionPart> Sections { get; } = [];

    public string AddImage(ImageElement image)
        => DocumentRelationships.AddImage(image);

    public string AddHyperlink(string url)
        => DocumentRelationships.AddHyperlink(url);

    public string AddChart(ChartElement chart)
        => DocumentRelationships.AddChart(chart);

    public RelationshipScope CreateRelationshipScope()
        => new(this, firstRelationshipId: 1);

    internal MediaEntry AddMedia(ImageElement image)
    {
        var key = ImageKey(image);
        if (_mediaByKey.TryGetValue(key, out var existing))
            return existing;

        var fileName = image.FilePath ?? image.FileName;
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(ext))
            ext = ".bin";
        var mediaPath = $"word/media/image{_nextMediaId++}{ext}";
        var entry = new MediaEntry(key, image.FilePath, image.Bytes, mediaPath, ImageReader.ContentType(fileName));
        _mediaByKey[key] = entry;
        return entry;
    }

    internal ChartPartEntry AddChartPart(ChartElement chart)
    {
        var entry = new ChartPartEntry($"word/charts/chart{_nextChartId++}.xml", chart);
        _charts.Add(entry);
        return entry;
    }

    public int NextDocPrId() => _nextDocPrId++;
    public int NextBookmarkId() => _nextBookmarkId++;
    public int GetNextFigureNumber() => _nextFigureNumber++;
    public int GetNextTableNumber() => _nextTableNumber++;

    public int AddFootnote(TextElement text)
    {
        var id = _nextFootnoteId++;
        Footnotes.Add(new NoteEntry(id, text));
        return id;
    }

    public int AddEndnote(TextElement text)
    {
        var id = _nextEndnoteId++;
        Endnotes.Add(new NoteEntry(id, text));
        return id;
    }

    public int AddList(ListElement list)
    {
        var key = NumberingKey(list);
        if (!_abstractNumberingByKey.TryGetValue(key, out var abstractDefinition))
        {
            abstractDefinition = new AbstractNumberingDefinition(
                _nextAbstractNumberId++,
                Enumerable.Range(0, 9)
                    .Select(level => NumberingLevel(list, level))
                    .ToList());
            _abstractNumberingByKey[key] = abstractDefinition;
        }

        var numberId = _nextNumberId++;
        _numberingInstances.Add(new NumberingInstance(numberId, abstractDefinition.AbstractNumberId));
        return numberId;
    }

    private static string ImageKey(ImageElement image)
    {
        if (!string.IsNullOrWhiteSpace(image.FilePath))
            return $"file:{Path.GetFullPath(image.FilePath)}";

        var hash = ComputeSha256Hex(image.Bytes ?? []);
        return $"bytes:{image.FileName}:{hash}";
    }

    private static string ComputeSha256Hex(byte[] bytes)
    {
        using var sha256 = SHA256.Create();
        return BitConverter.ToString(sha256.ComputeHash(bytes)).Replace("-", "");
    }

    private static string NumberingKey(ListElement list)
    {
        var markerKey = string.Join("|", list.Markers
            .OrderBy(x => x.Key)
            .Select(x => $"{x.Key}:{x.Value.Marker}:{x.Value.FontFamily}"));
        return $"{list.Ordered}:{markerKey}";
    }

    private static NumberingLevelDefinition NumberingLevel(ListElement list, int level)
    {
        var left = 720 + level * 360;
        const int hanging = 360;

        if (list.Markers.TryGetValue(level, out var marker))
        {
            return new NumberingLevelDefinition(
                level,
                list.Ordered ? "decimal" : "bullet",
                marker.Marker,
                left,
                hanging,
                marker.FontFamily);
        }

        if (list.Ordered)
        {
            return new NumberingLevelDefinition(
                level,
                "decimal",
                $"%{level + 1}.",
                left,
                hanging,
                null);
        }

        var defaults = (level % 3) switch
        {
            1 => ("\uF0A7", "Wingdings"),
            2 => ("\uF0B7", "Symbol"),
            _ => ("\uF0B7", "Symbol")
        };
        return new NumberingLevelDefinition(level, "bullet", defaults.Item1, left, hanging, defaults.Item2);
    }
}

internal sealed class RelationshipScope
{
    private readonly DocumentRenderContext _ctx;
    private readonly Dictionary<string, ImageEntry> _imagesByKey = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, HyperlinkEntry> _hyperlinksByUrl = new(StringComparer.OrdinalIgnoreCase);
    private int _nextRId;

    public RelationshipScope(DocumentRenderContext ctx, int firstRelationshipId)
    {
        _ctx = ctx;
        _nextRId = firstRelationshipId;
    }

    public List<ImageEntry> Images { get; } = [];
    public List<HyperlinkEntry> Hyperlinks { get; } = [];
    public List<ChartEntry> Charts { get; } = [];

    public string AddImage(ImageElement image)
    {
        var media = _ctx.AddMedia(image);
        if (_imagesByKey.TryGetValue(media.Key, out var existing))
            return existing.RelationshipId;

        var rId = NextRelationshipId();
        var entry = new ImageEntry(rId, media.MediaPath, media.ContentType);
        _imagesByKey[media.Key] = entry;
        Images.Add(entry);
        return rId;
    }

    public string AddHyperlink(string url)
    {
        if (_hyperlinksByUrl.TryGetValue(url, out var existing))
            return existing.RelationshipId;

        var rId = NextRelationshipId();
        var entry = new HyperlinkEntry(rId, url);
        _hyperlinksByUrl[url] = entry;
        Hyperlinks.Add(entry);
        return rId;
    }

    public string AddChart(ChartElement chart)
    {
        var part = _ctx.AddChartPart(chart);
        var rId = NextRelationshipId();
        Charts.Add(new ChartEntry(rId, part.ChartPath, chart));
        return rId;
    }

    public string NextRelationshipId() => $"rId{_nextRId++}";
}
