using System.IO.Compression;
using System.Globalization;
using System.Text;
using TerraFluent.Docx.Reporting.Core;
using TerraFluent.Docx.Reporting.Core.Elements;

namespace TerraFluent.Docx.Reporting.Ooxml;

internal static class OoxmlWriter
{
    public static void Write(DocumentContainerDescriptor doc, Stream output)
    {
        var ctx = new DocumentRenderContext();

        // Build section parts (header/footer XML) and register their rIds
        var sections = BuildSections(doc, ctx);

        // Build document body XML (registers image rIds as it encounters images)
        string documentXml = DocumentXml(doc, ctx, sections);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        WriteText(archive, "[Content_Types].xml", ContentTypesXml(ctx, sections));
        WriteText(archive, "_rels/.rels",          RelsXml());
        WriteText(archive, "docProps/core.xml",    CorePropsXml(doc));
        WriteText(archive, "docProps/app.xml",     AppPropsXml(doc));
        WriteText(archive, "word/_rels/document.xml.rels", DocumentRelsXml(ctx, sections));
        WriteText(archive, "word/styles.xml",      StylesXml(doc.ThemeSettings, doc.Styles));
        WriteText(archive, "word/settings.xml",    SettingsXml(sections));
        WriteText(archive, "word/numbering.xml",   NumberingXml(ctx));
        WriteText(archive, "word/document.xml",    documentXml);
        if (ctx.Footnotes.Count > 0)
        {
            WriteText(archive, "word/footnotes.xml", NotesXml("footnotes", "footnote", ctx.Footnotes, ctx.FootnoteRelationships));
            if (ctx.FootnoteRelationships.Images.Count > 0 || ctx.FootnoteRelationships.Hyperlinks.Count > 0 || ctx.FootnoteRelationships.Charts.Count > 0)
                WriteText(archive, RelationshipPartPath("word/footnotes.xml"), PartRelsXml(ctx.FootnoteRelationships.Images, ctx.FootnoteRelationships.Hyperlinks, ctx.FootnoteRelationships.Charts));
        }
        if (ctx.Endnotes.Count > 0)
        {
            WriteText(archive, "word/endnotes.xml", NotesXml("endnotes", "endnote", ctx.Endnotes, ctx.EndnoteRelationships));
            if (ctx.EndnoteRelationships.Images.Count > 0 || ctx.EndnoteRelationships.Hyperlinks.Count > 0 || ctx.EndnoteRelationships.Charts.Count > 0)
                WriteText(archive, RelationshipPartPath("word/endnotes.xml"), PartRelsXml(ctx.EndnoteRelationships.Images, ctx.EndnoteRelationships.Hyperlinks, ctx.EndnoteRelationships.Charts));
        }

        foreach (var chart in ctx.ChartParts)
            WriteText(archive, chart.ChartPath, ChartXml(chart.Chart));

        foreach (var sec in sections)
        {
            foreach (var part in sec.HeaderFooterParts)
            {
                WriteText(archive, part.PartName, part.Xml);
                if (part.Images.Count > 0 || part.Hyperlinks.Count > 0 || part.Charts.Count > 0)
                    WriteText(archive, RelationshipPartPath(part.PartName), PartRelsXml(part.Images, part.Hyperlinks, part.Charts));
            }
        }

        foreach (var media in ctx.Media)
            WriteMedia(archive, media);
    }

    // -------------------------------------------------------------------------
    // Section (header / footer) assembly
    // -------------------------------------------------------------------------

    private static List<SectionPart> BuildSections(DocumentContainerDescriptor doc, DocumentRenderContext ctx)
    {
        var parts = new List<SectionPart>();
        int headerIdx = 1;
        int footerIdx = 1;
        var lastHasContent = new Dictionary<(bool IsHeader, string ReferenceType), bool>();
        foreach (var page in doc.Pages)
        {
            var sectionParts = new List<HeaderFooterPart>();
            var hasWatermark = !string.IsNullOrWhiteSpace(page.WatermarkText);
            var hasFirstPageHeaderFooter = page.FirstPageHeaderContainer.Element.Elements.Count > 0 || page.FirstPageFooterContainer.Element.Elements.Count > 0;
            var hasEvenPageHeaderFooter = page.EvenPageHeaderContainer.Element.Elements.Count > 0 || page.EvenPageFooterContainer.Element.Elements.Count > 0;

            AddHeaderFooterPart(sectionParts, page.HeaderContainer.Element, isHeader: true, referenceType: "default", ref headerIdx, ctx, page.DefaultTextStyle, page, lastHasContent);
            if (hasFirstPageHeaderFooter)
            {
                AddHeaderFooterPart(sectionParts, page.FirstPageHeaderContainer.Element, isHeader: true, referenceType: "first", ref headerIdx, ctx, page.DefaultTextStyle, hasWatermark ? page : null, lastHasContent);
                AddHeaderFooterPart(sectionParts, page.FirstPageFooterContainer.Element, isHeader: false, referenceType: "first", ref footerIdx, ctx, page.DefaultTextStyle, null, lastHasContent);
            }
            AddHeaderFooterPart(sectionParts, page.EvenPageHeaderContainer.Element, isHeader: true, referenceType: "even", ref headerIdx, ctx, page.DefaultTextStyle, hasWatermark && hasEvenPageHeaderFooter ? page : null, lastHasContent);
            AddHeaderFooterPart(sectionParts, page.FooterContainer.Element, isHeader: false, referenceType: "default", ref footerIdx, ctx, page.DefaultTextStyle, null, lastHasContent);
            AddHeaderFooterPart(sectionParts, page.EvenPageFooterContainer.Element, isHeader: false, referenceType: "even", ref footerIdx, ctx, page.DefaultTextStyle, null, lastHasContent);

            parts.Add(new SectionPart(sectionParts));
        }
        return parts;
    }

    private static void AddHeaderFooterPart(
        List<HeaderFooterPart> parts,
        ContainerElement container,
        bool isHeader,
        string referenceType,
        ref int idx,
        DocumentRenderContext ctx,
        TextElement defaultTextStyle,
        PageDescriptor? watermarkPage,
        Dictionary<(bool IsHeader, string ReferenceType), bool> lastHasContent)
    {
        var hasContent = container.Elements.Count > 0 || !string.IsNullOrWhiteSpace(watermarkPage?.WatermarkText);
        var key = (isHeader, referenceType);

        // Per OOXML rules, a section that omits a header/footer reference of a given type
        // inherits that reference from the previous section. If a previous section defined
        // one but this section has nothing of its own, emit an explicit empty override so
        // that content (including watermarks) doesn't bleed into this section.
        if (!hasContent && !(lastHasContent.TryGetValue(key, out var hadContent) && hadContent))
            return;

        var relationships = ctx.CreateRelationshipScope();
        var relId = ctx.DocumentRelationships.NextRelationshipId();
        var prefix = isHeader ? "header" : "footer";
        var tag = isHeader ? "hdr" : "ftr";
        var partName = $"word/{prefix}{idx++}.xml";
        var xml = ZoneXml(tag, container, ctx, relationships, defaultTextStyle, watermarkPage);

        parts.Add(new HeaderFooterPart(
            isHeader,
            referenceType,
            relId,
            partName,
            xml,
            relationships.Images,
            relationships.Hyperlinks,
            relationships.Charts));

        lastHasContent[key] = hasContent;
    }

    private static string ZoneXml(string tag, ContainerElement container, DocumentRenderContext ctx, RelationshipScope relationships, TextElement defaultTextStyle, PageDescriptor? watermarkPage = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine($"""<w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word">""");
        if (!string.IsNullOrWhiteSpace(watermarkPage?.WatermarkText))
            sb.AppendLine(WatermarkShapeXml(watermarkPage));
        new BodyBuilder(sb, ctx, relationships, defaultTextStyle).WriteContainer(container);
        sb.AppendLine($"</w:{tag}>");
        return sb.ToString();
    }

    // -------------------------------------------------------------------------
    // document.xml
    // -------------------------------------------------------------------------

    private static string DocumentXml(DocumentContainerDescriptor doc, DocumentRenderContext ctx, List<SectionPart> sections)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">""");
        var background = doc.Pages.Select(p => p.BackgroundColor).FirstOrDefault(c => !string.IsNullOrWhiteSpace(c));
        if (!string.IsNullOrWhiteSpace(background))
            sb.AppendLine(BackgroundXml(background));
        sb.AppendLine("<w:body>");

        for (int i = 0; i < doc.Pages.Count; i++)
        {
            var page = doc.Pages[i];
            var sec  = sections[i];
            var bodyBuilder = new BodyBuilder(sb, ctx, ctx.DocumentRelationships, page.DefaultTextStyle);

            if (i > 0)
                sb.AppendLine(SectionBreakParagraph(doc.Pages[i - 1], sections[i - 1]));

            bodyBuilder.WriteContainer(page.ContentContainer.Element);
        }

        // Final sectPr (last/only page)
        if (doc.Pages.Count > 0)
            sb.AppendLine(SectPr(doc.Pages[doc.Pages.Count - 1], sections[sections.Count - 1]));

        sb.AppendLine("</w:body>");
        sb.AppendLine("</w:document>");
        return sb.ToString();
    }

    private static string SectionBreakParagraph(PageDescriptor page, SectionPart sec) =>
        $"<w:p><w:pPr><w:sectPr>{SectPrContent(page, sec)}</w:sectPr></w:pPr></w:p>";

    private static string SectPr(PageDescriptor page, SectionPart sec) =>
        $"<w:sectPr>{SectPrContent(page, sec)}</w:sectPr>";

    private static string SectPrContent(PageDescriptor page, SectionPart sec) =>
        $"{HdrFtrRefs(sec)}{PageSizeXml(page)}{MarginsXml(page)}{PageNumberXml(page)}{ColumnsXml(page)}{TitlePageXml(sec)}";

    private static string HdrFtrRefs(SectionPart sec)
    {
        var sb = new StringBuilder();
        foreach (var part in sec.HeaderFooterParts)
        {
            var tag = part.IsHeader ? "headerReference" : "footerReference";
            sb.Append($"""<w:{tag} w:type="{part.ReferenceType}" r:id="{part.RelationshipId}"/>""");
        }
        return sb.ToString();
    }

    private static string PageNumberXml(PageDescriptor page)
    {
        if (page.PageNumberStartValue == null && page.PageNumberFormat == null)
            return string.Empty;

        var attrs = new StringBuilder();
        if (page.PageNumberStartValue.HasValue)
            attrs.Append($" w:start=\"{page.PageNumberStartValue.Value}\"");
        if (!string.IsNullOrEmpty(page.PageNumberFormat))
            attrs.Append($" w:fmt=\"{page.PageNumberFormat}\"");

        return $"<w:pgNumType{attrs}/>";
    }

    private static string ColumnsXml(PageDescriptor page)
    {
        if (page.ColumnCount <= 1 && !page.ColumnSeparatorLine)
            return string.Empty;

        var spacing = Math.Max(0, (int)(page.ColumnSpacing * 20));
        var separator = page.ColumnSeparatorLine ? " w:sep=\"1\"" : string.Empty;
        return $"""<w:cols w:num="{page.ColumnCount}" w:space="{spacing}"{separator}/>""";
    }

    private static string TitlePageXml(SectionPart sec) =>
        sec.HeaderFooterParts.Any(p => p.ReferenceType == "first")
            ? "<w:titlePg/>"
            : string.Empty;

    private static string PageSizeXml(PageDescriptor page)
    {
        int w = (int)(page.Size.Width  * 20);
        int h = (int)(page.Size.Height * 20);
        return $"""<w:pgSz w:w="{w}" w:h="{h}"/>""";
    }

    private static string MarginsXml(PageDescriptor page)
    {
        int top    = (int)(page.MarginTop    * 20);
        int right  = (int)(page.MarginRight  * 20);
        int bottom = (int)(page.MarginBottom * 20);
        int left   = (int)(page.MarginLeft   * 20);
        return $"""<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}" w:header="709" w:footer="709" w:gutter="0"/>""";
    }

    // -------------------------------------------------------------------------
    // Static XML parts
    // -------------------------------------------------------------------------

    private static string ContentTypesXml(DocumentRenderContext ctx, List<SectionPart> sections)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""");
        sb.AppendLine("""  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""");
        sb.AppendLine("""  <Default Extension="xml"  ContentType="application/xml"/>""");
        sb.AppendLine("""  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>""");
        sb.AppendLine("""  <Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>""");
        sb.AppendLine("""  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>""");
        sb.AppendLine("""  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>""");
        if (ctx.Footnotes.Count > 0)
            sb.AppendLine("""  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>""");
        if (ctx.Endnotes.Count > 0)
            sb.AppendLine("""  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>""");
        foreach (var chart in ctx.ChartParts)
            sb.AppendLine($"""  <Override PartName="/{chart.ChartPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>""");
        sb.AppendLine("""  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>""");
        sb.AppendLine("""  <Override PartName="/docProps/app.xml"  ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>""");

        foreach (var sec in sections)
        {
            foreach (var part in sec.HeaderFooterParts)
            {
                var contentType = part.IsHeader
                    ? "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
                    : "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
                sb.AppendLine($"""  <Override PartName="/{part.PartName}" ContentType="{contentType}"/>""");
            }
        }

        // Unique image content types
        foreach (var ct in ctx.Media.Select(i => i.ContentType).Distinct())
        {
            var ext = ct == "image/png" ? "png" : ct == "image/jpeg" ? "jpg" : ct == "image/gif" ? "gif" : "bin";
            sb.AppendLine($"""  <Default Extension="{ext}" ContentType="{ct}"/>""");
            if (ct == "image/jpeg")
                sb.AppendLine($"""  <Default Extension="jpeg" ContentType="{ct}"/>""");
        }

        sb.AppendLine("</Types>");
        return sb.ToString();
    }

    private static string PartRelsXml(IReadOnlyList<ImageEntry> images, IReadOnlyList<HyperlinkEntry> hyperlinks, IReadOnlyList<ChartEntry> charts)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""");
        foreach (var img in images)
            sb.AppendLine($"""  <Relationship Id="{img.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{img.MediaPath.Replace("word/", "")}"/>""");
        foreach (var link in hyperlinks)
            sb.AppendLine($"""  <Relationship Id="{link.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{Escape(link.Url)}" TargetMode="External"/>""");
        foreach (var chart in charts)
            sb.AppendLine($"""  <Relationship Id="{chart.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{chart.ChartPath.Replace("word/", "")}"/>""");
        sb.AppendLine("</Relationships>");
        return sb.ToString();
    }

    private static string RelationshipPartPath(string partName)
    {
        var directory = Path.GetDirectoryName(partName)?.Replace('\\', '/');
        var fileName = Path.GetFileName(partName);
        return string.IsNullOrEmpty(directory)
            ? $"_rels/{fileName}.rels"
            : $"{directory}/_rels/{fileName}.rels";
    }

    private static string RelsXml() => """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        </Relationships>
        """;

    private static string DocumentRelsXml(DocumentRenderContext ctx, List<SectionPart> sections)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""");
        sb.AppendLine("""  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"   Target="styles.xml"/>""");
        sb.AppendLine("""  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>""");
        sb.AppendLine("""  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>""");
        if (ctx.Footnotes.Count > 0)
            sb.AppendLine("""  <Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>""");
        if (ctx.Endnotes.Count > 0)
            sb.AppendLine("""  <Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>""");

        foreach (var sec in sections)
        {
            foreach (var part in sec.HeaderFooterParts)
            {
                var type = part.IsHeader ? "header" : "footer";
                sb.AppendLine($"""  <Relationship Id="{part.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{type}" Target="{Path.GetFileName(part.PartName)}"/>""");
            }
        }

        foreach (var img in ctx.Images)
            sb.AppendLine($"""  <Relationship Id="{img.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{img.MediaPath.Replace("word/", "")}"/>""");
        foreach (var link in ctx.Hyperlinks)
            sb.AppendLine($"""  <Relationship Id="{link.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{Escape(link.Url)}" TargetMode="External"/>""");
        foreach (var chart in ctx.Charts)
            sb.AppendLine($"""  <Relationship Id="{chart.RelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{chart.ChartPath.Replace("word/", "")}"/>""");

        sb.AppendLine("</Relationships>");
        return sb.ToString();
    }

    private static string SettingsXml(List<SectionPart> sections)
    {
        var evenOdd = sections.SelectMany(s => s.HeaderFooterParts).Any(p => p.ReferenceType == "even")
            ? "  <w:evenAndOddHeaders/>\n"
            : string.Empty;
        return $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:displayBackgroundShape/>
          <w:defaultTabStop w:val="720"/>
        {evenOdd}</w:settings>
        """;
    }

    private static string NumberingXml(DocumentRenderContext ctx)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""");

        foreach (var definition in ctx.AbstractNumberingDefinitions.OrderBy(d => d.AbstractNumberId))
        {
            sb.AppendLine($"""  <w:abstractNum w:abstractNumId="{definition.AbstractNumberId}">""");
            sb.AppendLine("""    <w:multiLevelType w:val="hybridMultilevel"/>""");
            foreach (var level in definition.Levels)
            {
                sb.AppendLine($"""    <w:lvl w:ilvl="{level.Level}">""");
                sb.AppendLine("""      <w:start w:val="1"/>""");
                sb.AppendLine($"""      <w:numFmt w:val="{level.NumberFormat}"/>""");
                sb.AppendLine($"""      <w:lvlText w:val="{Escape(level.LevelText)}"/>""");
                sb.AppendLine("""      <w:lvlJc w:val="left"/>""");
                sb.AppendLine($"""      <w:pPr><w:ind w:left="{level.LeftIndent}" w:hanging="{level.HangingIndent}"/></w:pPr>""");
                if (!string.IsNullOrWhiteSpace(level.FontFamily))
                    sb.AppendLine($"""      <w:rPr><w:rFonts w:ascii="{Escape(level.FontFamily)}" w:hAnsi="{Escape(level.FontFamily)}" w:hint="default"/></w:rPr>""");
                sb.AppendLine("""    </w:lvl>""");
            }
            sb.AppendLine("""  </w:abstractNum>""");
        }

        foreach (var instance in ctx.NumberingInstances.OrderBy(i => i.NumberId))
            sb.AppendLine($"""  <w:num w:numId="{instance.NumberId}"><w:abstractNumId w:val="{instance.AbstractNumberId}"/></w:num>""");

        sb.AppendLine("</w:numbering>");
        return sb.ToString();
    }

    private static string NotesXml(string rootName, string itemName, IReadOnlyList<NoteEntry> notes, RelationshipScope relationships)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine($"""<w:{rootName} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""");
        sb.AppendLine($"""  <w:{itemName} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{itemName}>""");
        sb.AppendLine($"""  <w:{itemName} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{itemName}>""");
        foreach (var note in notes)
        {
            sb.AppendLine($"""  <w:{itemName} w:id="{note.Id}">""");
            AppendNoteParagraph(sb, note.Text, relationships);
            sb.AppendLine($"""  </w:{itemName}>""");
        }
        sb.AppendLine($"""</w:{rootName}>""");
        return sb.ToString();
    }

    private static void AppendNoteParagraph(StringBuilder sb, TextElement text, RelationshipScope relationships)
    {
        sb.Append("    <w:p>");
        var paragraphDefaultRun = text.Runs.Count > 0 ? text.Runs[0] : null;
        foreach (var run in text.Runs)
        {
            if (run.Kind == TextRunKind.Text && run.Text.Length == 0)
                continue;

            var rPr = BuildStyleRPr(run, paragraphDefaultRun);
            var rPrXml = rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : string.Empty;

            switch (run.Kind)
            {
                case TextRunKind.Tab:
                    sb.Append($"<w:r>{rPrXml}<w:tab/></w:r>");
                    break;
                case TextRunKind.Hyperlink:
                    AppendNoteHyperlinkRun(sb, run, rPrXml, relationships);
                    break;
                case TextRunKind.Field:
                    AppendNoteFieldRun(sb, rPrXml, run.FieldInstruction ?? string.Empty, run.FieldCachedText ?? string.Empty);
                    break;
                case TextRunKind.PageNumber:
                    AppendNoteFieldRun(sb, rPrXml, "PAGE", "1");
                    break;
                case TextRunKind.PageCount:
                    AppendNoteFieldRun(sb, rPrXml, "NUMPAGES", "1");
                    break;
                default:
                    sb.Append($"""<w:r>{rPrXml}<w:t xml:space="preserve">{Escape(run.Text)}</w:t></w:r>""");
                    break;
            }
        }
        sb.AppendLine("</w:p>");
    }

    private static void AppendNoteHyperlinkRun(StringBuilder sb, TextRun run, string rPrXml, RelationshipScope relationships)
    {
        if (string.IsNullOrWhiteSpace(run.Url))
            return;

        var rId = relationships.AddHyperlink(run.Url);
        sb.Append($"""<w:hyperlink r:id="{rId}" w:history="1"><w:r>{rPrXml}<w:t xml:space="preserve">{Escape(run.Text)}</w:t></w:r></w:hyperlink>""");
    }

    // A OOXML field run: begin → instrText → separate → cached value → end
    private static void AppendNoteFieldRun(StringBuilder sb, string rPrXml, string fieldInstruction, string cachedText)
    {
        sb.Append($"""<w:r>{rPrXml}<w:fldChar w:fldCharType="begin"/></w:r>""");
        sb.Append($"""<w:r>{rPrXml}<w:instrText xml:space="preserve"> {Escape(fieldInstruction)} </w:instrText></w:r>""");
        sb.Append($"""<w:r>{rPrXml}<w:fldChar w:fldCharType="separate"/></w:r>""");
        sb.Append($"""<w:r>{rPrXml}<w:t xml:space="preserve">{Escape(cachedText)}</w:t></w:r>""");
        sb.Append($"""<w:r>{rPrXml}<w:fldChar w:fldCharType="end"/></w:r>""");
    }

    private static string CorePropsXml(DocumentContainerDescriptor doc)
    {
        var now = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture);
        return $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <cp:coreProperties
          xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
          xmlns:dc="http://purl.org/dc/elements/1.1/"
          xmlns:dcterms="http://purl.org/dc/terms/"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
          <dc:title>{Escape(doc.Title)}</dc:title>
          <dc:creator>{Escape(doc.Author)}</dc:creator>
          <cp:lastModifiedBy>{Escape(doc.Creator ?? doc.Author)}</cp:lastModifiedBy>
          <dc:subject>{Escape(doc.Subject)}</dc:subject>
          <cp:keywords>{Escape(doc.Keywords)}</cp:keywords>
          <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
          <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
        </cp:coreProperties>
        """;
    }

    private static string AppPropsXml(DocumentContainerDescriptor doc) => $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
          <Application>{Escape(doc.Creator ?? "TerraFluent.Docx.Reporting")}</Application>
        </Properties>
        """;

    private static string StylesXml(DocumentTheme theme, DocumentStyleCatalog styles)
    {
        var font = Escape(theme.DefaultFontFamily);
        var defaultFontSize = HalfPoints(theme.DefaultFontSize);
        var defaultTextColor = theme.DefaultTextColor.TrimStart('#');
        var headingColor = theme.HeadingColor.TrimStart('#');
        var sb = new StringBuilder();
        sb.Append($"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults>
            <w:rPrDefault>
              <w:rPr>
                <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>
                <w:color w:val="{defaultTextColor}"/>
                <w:sz w:val="{defaultFontSize}"/><w:szCs w:val="{defaultFontSize}"/>
                <w:lang w:val="en-US"/>
              </w:rPr>
            </w:rPrDefault>
            <w:pPrDefault>
              <w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>
            </w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
          <w:style w:type="paragraph" w:styleId="Heading1">
            <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
            <w:rPr><w:b/><w:color w:val="{headingColor}"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading2">
            <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
            <w:rPr><w:b/><w:color w:val="{headingColor}"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading3">
            <w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
            <w:rPr><w:b/><w:color w:val="{headingColor}"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading4">
            <w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="3"/></w:pPr>
            <w:rPr><w:b/><w:color w:val="{headingColor}"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading5">
            <w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="4"/></w:pPr>
            <w:rPr><w:b/><w:color w:val="{headingColor}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading6">
            <w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
            <w:pPr><w:outlineLvl w:val="5"/></w:pPr>
            <w:rPr><w:color w:val="{headingColor}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
          </w:style>
        """);

        foreach (var style in styles.ParagraphStyles.OrderBy(s => s.StyleId))
            AppendParagraphStyle(sb, style);

        foreach (var style in styles.TableStyles.OrderBy(s => s.StyleId))
            AppendTableStyle(sb, style);

        sb.AppendLine("</w:styles>");
        return sb.ToString();
    }

    private static void AppendParagraphStyle(StringBuilder sb, ParagraphStyleDefinition style)
    {
        sb.AppendLine($"""  <w:style w:type="paragraph" w:customStyle="1" w:styleId="{Escape(style.StyleId)}">""");
        sb.AppendLine($"""    <w:name w:val="{Escape(style.Name)}"/>""");
        sb.AppendLine("""    <w:basedOn w:val="Normal"/><w:qFormat/>""");

        var pPr = BuildStylePPr(style.Style);
        if (pPr.Length > 0)
            sb.AppendLine($"    <w:pPr>{pPr}</w:pPr>");

        var run = style.Style.Runs.FirstOrDefault();
        var rPr = run != null ? BuildStyleRPr(run) : string.Empty;
        if (rPr.Length > 0)
            sb.AppendLine($"    <w:rPr>{rPr}</w:rPr>");

        sb.AppendLine("""  </w:style>""");
    }

    private static void AppendTableStyle(StringBuilder sb, TableStyleDefinition style)
    {
        sb.AppendLine($"""  <w:style w:type="table" w:customStyle="1" w:styleId="{Escape(style.StyleId)}">""");
        sb.AppendLine($"""    <w:name w:val="{Escape(style.Name)}"/>""");
        sb.AppendLine("""    <w:basedOn w:val="TableNormal"/><w:qFormat/>""");
        sb.AppendLine("""  </w:style>""");
    }

    private static string BuildStylePPr(TextElement text)
    {
        var sb = new StringBuilder();
        if (text.KeepWithNext)
            sb.Append("<w:keepNext/>");
        if (text.KeepLinesTogether)
            sb.Append("<w:keepLines/>");
        if (text.PageBreakBefore)
            sb.Append("<w:pageBreakBefore/>");
        AppendStyleParagraphBorders(sb, text);
        if (!string.IsNullOrWhiteSpace(text.ShadingColor))
            sb.Append($"""<w:shd w:val="clear" w:color="auto" w:fill="{text.ShadingColor.TrimStart('#')}"/>""");
        if (text.SpacingBefore > 0 || text.SpacingAfter > 0 || text.LineHeight.HasValue)
        {
            sb.Append("<w:spacing");
            sb.Append($" w:before=\"{ToDxa(text.SpacingBefore)}\" w:after=\"{ToDxa(text.SpacingAfter)}\"");
            if (text.LineHeight.HasValue)
                sb.Append($" w:line=\"{(int)(text.LineHeight.Value * 240)}\" w:lineRule=\"auto\"");
            sb.Append("/>");
        }
        if (text.LeftIndent.HasValue || text.RightIndent.HasValue || text.FirstLineIndent.HasValue || text.HangingIndent.HasValue)
        {
            sb.Append("<w:ind");
            if (text.LeftIndent.HasValue)
                sb.Append($" w:left=\"{ToDxa(text.LeftIndent.Value)}\"");
            if (text.RightIndent.HasValue)
                sb.Append($" w:right=\"{ToDxa(text.RightIndent.Value)}\"");
            if (text.FirstLineIndent.HasValue)
                sb.Append($" w:firstLine=\"{ToDxa(text.FirstLineIndent.Value)}\"");
            if (text.HangingIndent.HasValue)
                sb.Append($" w:hanging=\"{ToDxa(text.HangingIndent.Value)}\"");
            sb.Append("/>");
        }
        var jc = text.Alignment switch
        {
            "center" => "center",
            "right" => "right",
            "both" => "both",
            _ => null
        };
        if (jc != null)
            sb.Append($"""<w:jc w:val="{jc}"/>""");
        return sb.ToString();
    }

    private static void AppendStyleParagraphBorders(StringBuilder sb, TextElement text)
    {
        if (text.TopBorder == null && text.RightBorder == null && text.BottomBorder == null && text.LeftBorder == null)
            return;

        sb.Append("<w:pBdr>");
        AppendStyleParagraphBorder(sb, "top", text.TopBorder);
        AppendStyleParagraphBorder(sb, "left", text.LeftBorder);
        AppendStyleParagraphBorder(sb, "bottom", text.BottomBorder);
        AppendStyleParagraphBorder(sb, "right", text.RightBorder);
        sb.Append("</w:pBdr>");
    }

    private static void AppendStyleParagraphBorder(StringBuilder sb, string side, ParagraphBorder? border)
    {
        if (border == null)
            return;

        var size = Math.Max(1, (int)(border.Width * 8));
        var space = Math.Max(0, (int)border.Space);
        sb.Append($"""<w:{side} w:val="single" w:sz="{size}" w:space="{space}" w:color="{border.Color.TrimStart('#')}"/>""");
    }

    private static string BuildStyleRPr(TextRun run, TextRun? paragraphDefaultRun = null)
    {
        // Element order follows CT_RPr sequence: rFonts, b, i, strike, color, sz, szCs, highlight, u, vertAlign
        var sb = new StringBuilder();
        string? fontFamily = run.FontFamily ?? paragraphDefaultRun?.FontFamily;
        bool bold = run.Bold ?? paragraphDefaultRun?.Bold ?? false;
        bool italic = run.Italic ?? paragraphDefaultRun?.Italic ?? false;
        bool strikethrough = run.Strikethrough ?? paragraphDefaultRun?.Strikethrough ?? false;
        bool underline = run.Underline ?? paragraphDefaultRun?.Underline ?? false;
        string? fontColor = run.FontColor ?? paragraphDefaultRun?.FontColor;
        float? fontSize = run.FontSize ?? paragraphDefaultRun?.FontSize;
        string? highlightColor = run.HighlightColor ?? paragraphDefaultRun?.HighlightColor;
        string? verticalAlignment = run.VerticalAlignment ?? paragraphDefaultRun?.VerticalAlignment;

        if (fontFamily != null)
            sb.Append($"""<w:rFonts w:ascii="{Escape(fontFamily)}" w:hAnsi="{Escape(fontFamily)}"/>""");
        if (bold)
            sb.Append("<w:b/>");
        if (italic)
            sb.Append("<w:i/>");
        if (strikethrough)
            sb.Append("<w:strike/>");
        if (fontColor != null)
            sb.Append($"""<w:color w:val="{fontColor.TrimStart('#')}"/>""");
        if (fontSize.HasValue)
        {
            var hp = HalfPoints(fontSize.Value);
            sb.Append($"""<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>""");
        }
        if (highlightColor != null)
            sb.Append($"""<w:highlight w:val="{highlightColor}"/>""");
        if (underline)
            sb.Append("<w:u w:val=\"single\"/>");
        if (verticalAlignment != null)
            sb.Append($"""<w:vertAlign w:val="{verticalAlignment}"/>""");
        return sb.ToString();
    }

    private static int ToDxa(float points) => Math.Max(0, (int)(points * 20));

    private static string BackgroundXml(string color)
    {
        var value = color.TrimStart('#');
        return $"""
        <w:background w:color="{value}">
          <v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">
            <v:fill color2="#{value}" type="solid"/>
          </v:background>
        </w:background>
        """;
    }

    private static string WatermarkShapeXml(PageDescriptor page)
    {
        var text = Escape(page.WatermarkText);
        var color = page.WatermarkColor.TrimStart('#');
        var fontSize = Math.Max(72, page.WatermarkFontSize);
        var width = MathCompat.Clamp((page.WatermarkText?.Length ?? 0) * fontSize * 0.9, 560, 820);
        var height = MathCompat.Clamp(fontSize * 2.0, 140, 220);
        var widthText = width.ToString("0.##", CultureInfo.InvariantCulture);
        var heightText = height.ToString("0.##", CultureInfo.InvariantCulture);

        return $"""
        <w:p>
          <w:r>
            <w:pict>
              <v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
                <v:formulas>
                  <v:f eqn="sum #0 0 10800"/>
                  <v:f eqn="prod #0 2 1"/>
                  <v:f eqn="sum 21600 0 @1"/>
                  <v:f eqn="sum 0 0 @2"/>
                  <v:f eqn="sum 21600 0 @3"/>
                  <v:f eqn="if @0 @3 0"/>
                  <v:f eqn="if @0 21600 @1"/>
                  <v:f eqn="if @0 0 @2"/>
                  <v:f eqn="if @0 @4 21600"/>
                  <v:f eqn="mid @5 @6"/>
                  <v:f eqn="mid @8 @5"/>
                  <v:f eqn="mid @7 @8"/>
                  <v:f eqn="mid @6 @7"/>
                  <v:f eqn="sum @6 0 @5"/>
                </v:formulas>
                <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>
                <v:textpath on="t" fitshape="t"/>
                <v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>
                <o:lock v:ext="edit" text="t" shapetype="t"/>
              </v:shapetype>
              <v:shape id="TerraFluent.Docx.ReportingWatermark" o:spid="_x0000_s1026" type="#_x0000_t136"
                style="position:absolute;margin-left:0;margin-top:0;width:{widthText}pt;height:{heightText}pt;rotation:315;z-index:-251654144;visibility:visible;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"
                o:allowincell="f" fillcolor="#{color}" stroked="f">
                <v:fill opacity=".55"/>
                <v:textpath style="font-family:&quot;Aptos&quot;;font-size:{fontSize.ToString("0.##", CultureInfo.InvariantCulture)}pt" string="{text}"/>
                <w10:wrap anchorx="margin" anchory="margin"/>
              </v:shape>
            </w:pict>
          </w:r>
        </w:p>
        """;
    }

    private static string ChartXml(ChartElement chart)
    {
        if (chart.Series.Count == 0)
            chart.Series.Add(new ChartSeriesElement());

        // Use the first series that actually has data points to determine the chart type,
        // since a leading empty series would otherwise report its unused default "bar" kind.
        var firstSeries = chart.Series.FirstOrDefault(s => s.Points.Count > 0) ?? chart.Series[0];
        var kind = firstSeries.Kind.ToLowerInvariant();
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">""");
        sb.AppendLine("""  <c:date1904 val="0"/>""");
        sb.AppendLine("""  <c:lang val="en-US"/>""");
        sb.AppendLine("""  <c:roundedCorners val="0"/>""");
        sb.AppendLine("""  <c:chart>""");
        sb.AppendLine(ChartTitleXml(chart.Title));
        sb.AppendLine("""    <c:autoTitleDeleted val="0"/>""");
        sb.AppendLine("""    <c:plotArea>""");
        sb.AppendLine("""      <c:layout/>""");
        sb.AppendLine(kind switch
        {
            "line" => LineChartXml(chart),
            "pie" => PieChartXml(chart, doughnut: false),
            "doughnut" => PieChartXml(chart, doughnut: true),
            _ => BarChartXml(chart)
        });
        if (kind is not "pie" and not "doughnut")
        {
            sb.AppendLine(CategoryAxisXml());
            sb.AppendLine(ValueAxisXml());
        }
        sb.AppendLine("""    </c:plotArea>""");
        sb.AppendLine("""    <c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend>""");
        sb.AppendLine("""    <c:plotVisOnly val="1"/>""");
        sb.AppendLine("""    <c:dispBlanksAs val="gap"/>""");
        sb.AppendLine("""    <c:showDLblsOverMax val="0"/>""");
        sb.AppendLine("""  </c:chart>""");
        sb.AppendLine("""
          <c:printSettings>
            <c:headerFooter/>
            <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
            <c:pageSetup/>
          </c:printSettings>
        """);
        sb.AppendLine("</c:chartSpace>");
        return sb.ToString();
    }

    private static string ChartTitleXml(string title) => $"""
            <c:title>
              <c:tx>
                <c:rich>
                  <a:bodyPr/>
                  <a:lstStyle/>
                  <a:p>
                    <a:r><a:rPr lang="en-US" sz="1400"/><a:t>{Escape(title)}</a:t></a:r>
                    <a:endParaRPr lang="en-US"/>
                  </a:p>
                </c:rich>
              </c:tx>
              <c:layout/>
              <c:overlay val="0"/>
            </c:title>
        """;

    private static string BarChartXml(ChartElement chart)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""              <c:barChart>""");
        sb.AppendLine("""                <c:barDir val="col"/>""");
        sb.AppendLine("""                <c:grouping val="clustered"/>""");
        sb.AppendLine("""                <c:varyColors val="0"/>""");

        int seriesIndex = 0;
        foreach (var series in chart.Series)
        {
            sb.Append(SeriesXml(series, seriesIndex, "bar"));
            seriesIndex++;
        }

        sb.AppendLine("""                <c:gapWidth val="150"/>""");
        sb.AppendLine("""                <c:axId val="48650112"/>""");
        sb.AppendLine("""                <c:axId val="48672768"/>""");
        sb.AppendLine("""              </c:barChart>""");
        return sb.ToString();
    }

    private static string LineChartXml(ChartElement chart)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""              <c:lineChart>""");
        sb.AppendLine("""                <c:grouping val="standard"/>""");
        sb.AppendLine("""                <c:varyColors val="0"/>""");

        int seriesIndex = 0;
        foreach (var series in chart.Series)
        {
            sb.Append(SeriesXmlWithMarker(series, seriesIndex));
            seriesIndex++;
        }

        sb.AppendLine("""                <c:marker val="1"/>""");
        sb.AppendLine("""                <c:smooth val="0"/>""");
        sb.AppendLine("""                <c:axId val="48650112"/>""");
        sb.AppendLine("""                <c:axId val="48672768"/>""");
        sb.AppendLine("""              </c:lineChart>""");
        return sb.ToString();
    }

    private static string PieChartXml(ChartElement chart, bool doughnut)
    {
        // Pie charts support only single series
        if (chart.Series.Count > 1)
            throw new InvalidOperationException("Pie and doughnut charts support only a single data series");

        var tag = doughnut ? "doughnutChart" : "pieChart";
        var holeSize = doughnut ? """<c:holeSize val="55"/>""" : string.Empty;
        var series = chart.Series.FirstOrDefault() ?? new ChartSeriesElement();
        var sb = new StringBuilder();
        sb.AppendLine($"""              <c:{tag}>""");
        sb.AppendLine("""                <c:varyColors val="1"/>""");
        sb.Append(SeriesXml(series, 0, doughnut ? "doughnut" : "pie"));
        sb.AppendLine("""                <c:firstSliceAng val="0"/>""");
        sb.AppendLine(holeSize);
        sb.AppendLine($"""              </c:{tag}>""");
        return sb.ToString();
    }

    private static string SeriesXml(ChartSeriesElement series, int index, string chartKind)
    {
        var color = ChartColor(series.Color);
        var sb = new StringBuilder();
        sb.AppendLine($"""                <c:ser>""");
        sb.AppendLine($"""                  <c:idx val="{index}"/>""");
        sb.AppendLine($"""                  <c:order val="{index}"/>""");
        if (!string.IsNullOrEmpty(series.Name))
            sb.AppendLine($"""                  <c:tx><c:v>{Escape(series.Name)}</c:v></c:tx>""");

        if (chartKind == "bar")
            sb.Append(SeriesShapeXml(color));
        else if (chartKind == "line")
            sb.Append(LineShapeXml(color));

        sb.Append(CategoriesXml(series));
        sb.Append(ValuesXml(series));
        sb.AppendLine("""                </c:ser>""");
        return sb.ToString();
    }

    private static string SeriesXmlWithMarker(ChartSeriesElement series, int index)
    {
        var color = ChartColor(series.Color);
        var sb = new StringBuilder();
        sb.AppendLine($"""                <c:ser>""");
        sb.AppendLine($"""                  <c:idx val="{index}"/>""");
        sb.AppendLine($"""                  <c:order val="{index}"/>""");
        if (!string.IsNullOrEmpty(series.Name))
            sb.AppendLine($"""                  <c:tx><c:v>{Escape(series.Name)}</c:v></c:tx>""");

        sb.Append(LineShapeXml(color));
        sb.AppendLine("""                  <c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>""");
        sb.Append(CategoriesXml(series));
        sb.Append(ValuesXml(series));
        sb.AppendLine("""                  <c:smooth val="0"/>""");
        sb.AppendLine("""                </c:ser>""");
        return sb.ToString();
    }

    private static string CategoriesXml(ChartSeriesElement series)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""                  <c:cat><c:strLit>""");
        sb.AppendLine($"""                    <c:ptCount val="{series.Points.Count}"/>""");
        for (int i = 0; i < series.Points.Count; i++)
            sb.AppendLine($"""                    <c:pt idx="{i}"><c:v>{Escape(series.Points[i].Label)}</c:v></c:pt>""");
        sb.AppendLine("""                  </c:strLit></c:cat>""");
        return sb.ToString();
    }

    private static string ValuesXml(ChartSeriesElement series)
    {
        var sb = new StringBuilder();
        sb.AppendLine("""                  <c:val><c:numLit>""");
        sb.AppendLine("""                    <c:formatCode>General</c:formatCode>""");
        sb.AppendLine($"""                    <c:ptCount val="{series.Points.Count}"/>""");
        for (int i = 0; i < series.Points.Count; i++)
            sb.AppendLine($"""                    <c:pt idx="{i}"><c:v>{ChartValue(series.Points[i].Value)}</c:v></c:pt>""");
        sb.AppendLine("""                  </c:numLit></c:val>""");
        return sb.ToString();
    }

    private static string SeriesShapeXml(string color) => $"""
                  <c:spPr>
                    <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                    <a:ln><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
                  </c:spPr>
        """;

    private static string LineShapeXml(string color) => $"""
                  <c:spPr>
                    <a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
                  </c:spPr>
        """;

    private static string CategoryAxisXml() => """
              <c:catAx>
                <c:axId val="48650112"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="b"/>
                <c:majorTickMark val="none"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="48672768"/>
                <c:crosses val="autoZero"/>
                <c:auto val="1"/>
                <c:lblAlgn val="ctr"/>
                <c:lblOffset val="100"/>
                <c:noMultiLvlLbl val="0"/>
              </c:catAx>
        """;

    private static string ValueAxisXml() => """
              <c:valAx>
                <c:axId val="48672768"/>
                <c:scaling><c:orientation val="minMax"/></c:scaling>
                <c:delete val="0"/>
                <c:axPos val="l"/>
                <c:majorGridlines/>
                <c:numFmt formatCode="General" sourceLinked="1"/>
                <c:majorTickMark val="none"/>
                <c:minorTickMark val="none"/>
                <c:tickLblPos val="nextTo"/>
                <c:crossAx val="48650112"/>
                <c:crosses val="autoZero"/>
                <c:crossBetween val="between"/>
              </c:valAx>
        """;

    private static string ChartValue(double value) =>
        value.ToString("0.############", CultureInfo.InvariantCulture);

    private static string ChartColor(string color) =>
        string.IsNullOrWhiteSpace(color)
            ? "1976D2"
            : color.TrimStart('#');

    // -------------------------------------------------------------------------
    // ZIP helpers
    // -------------------------------------------------------------------------

    private static void WriteText(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new System.Text.UTF8Encoding(false));
        writer.Write(content);
    }

    private static void WriteMedia(ZipArchive archive, MediaEntry media)
    {
        var entry = archive.CreateEntry(media.MediaPath, CompressionLevel.Optimal);
        using var dest = entry.Open();
        if (media.Bytes != null)
        {
            dest.Write(media.Bytes, 0, media.Bytes.Length);
            return;
        }

        using var src = File.OpenRead(media.FilePath!);
        src.CopyTo(dest);
    }

    internal static string Escape(string? s) =>
        s is null ? string.Empty : s
            .Replace("&",  "&amp;")
            .Replace("<",  "&lt;")
            .Replace(">",  "&gt;")
            .Replace("\"", "&quot;");

    private static int HalfPoints(float points) => Math.Max(1, (int)(points * 2));
}
