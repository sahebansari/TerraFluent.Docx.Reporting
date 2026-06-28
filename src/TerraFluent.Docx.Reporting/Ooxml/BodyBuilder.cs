using System.Text;
using TerraFluent.Docx.Reporting.Core;
using TerraFluent.Docx.Reporting.Core.Elements;

namespace TerraFluent.Docx.Reporting.Ooxml;

internal sealed class BodyBuilder
{
    private readonly StringBuilder _sb;
    private readonly DocumentRenderContext _ctx;
    private readonly RelationshipScope _relationships;
    private readonly TextElement? _defaultTextStyle;
    private readonly TextRun? _defaultRun;

    public BodyBuilder(StringBuilder sb, DocumentRenderContext ctx, RelationshipScope relationships, TextElement? defaultTextStyle = null)
    {
        _sb  = sb;
        _ctx = ctx;
        _relationships = relationships;
        _defaultTextStyle = defaultTextStyle;
        _defaultRun = defaultTextStyle?.Runs.FirstOrDefault();
    }

    public void WriteContainer(ContainerElement container)
    {
        foreach (var el in container.Elements)
            WriteElement(el);
    }

    private void WriteElement(IElement el)
    {
        switch (el)
        {
            case TextElement    text:  WriteParagraph(text);        break;
            case ColumnElement  col:   WriteColumn(col);            break;
            case RowElement     row:   WriteRow(row);               break;
            case ListElement    list:  WriteList(list);             break;
            case TableElement   table: WriteTable(table);           break;
            case BookmarkElement bookmark: WriteBookmark(bookmark);  break;
            case TableOfContentsElement toc: WriteTableOfContents(toc); break;
            case ChartElement chart: WriteChart(chart);             break;
            case ImageElement   img:   WriteImage(img);             break;
            case LineElement:          WriteHorizontalLine();       break;
            case PageBreakElement:     WritePageBreak();            break;
        }
    }

    // -------------------------------------------------------------------------
    // Paragraph
    // -------------------------------------------------------------------------

    private void WriteParagraph(TextElement text, int? numId = null, int listLevel = 0)
    {
        _sb.Append("<w:p>");
        bool useDefaultTextStyle = text.StyleId == null;

        // Build pPr content; only emit the element when it has children
        var pPr = new StringBuilder();

        if (text.StyleId != null)
            pPr.Append($"<w:pStyle w:val=\"{text.StyleId}\"/>");

        if (text.KeepWithNext)
            pPr.Append("<w:keepNext/>");
        if (text.KeepLinesTogether)
            pPr.Append("<w:keepLines/>");
        if (text.PageBreakBefore)
            pPr.Append("<w:pageBreakBefore/>");

        if (numId.HasValue)
            pPr.Append($"<w:numPr><w:ilvl w:val=\"{MathCompat.Clamp(listLevel, 0, 8)}\"/><w:numId w:val=\"{numId.Value}\"/></w:numPr>");

        WriteParagraphBorders(text, pPr);

        if (!string.IsNullOrWhiteSpace(text.ShadingColor))
            pPr.Append($"<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{text.ShadingColor.TrimStart('#')}\"/>");

        if (text.TabStops.Count > 0)
        {
            pPr.Append("<w:tabs>");
            foreach (var tab in text.TabStops)
                pPr.Append($"<w:tab w:val=\"{TabStopAlignment(tab.Alignment)}\" w:pos=\"{ToDxa(tab.Position)}\"/>");
            pPr.Append("</w:tabs>");
        }

        // spacing must precede jc in CT_PPrBase sequence
        float spacingBefore = text.SpacingBefore > 0 ? text.SpacingBefore : useDefaultTextStyle ? _defaultTextStyle?.SpacingBefore ?? 0 : 0;
        float spacingAfter = text.SpacingAfter > 0 ? text.SpacingAfter : useDefaultTextStyle ? _defaultTextStyle?.SpacingAfter ?? 0 : 0;
        float? lineHeight = text.LineHeight ?? (useDefaultTextStyle ? _defaultTextStyle?.LineHeight : null);

        if (spacingBefore > 0 || spacingAfter > 0 || lineHeight.HasValue)
        {
            int before = (int)(spacingBefore * 20);
            int after  = (int)(spacingAfter  * 20);
            if (lineHeight.HasValue)
            {
                int line = (int)(lineHeight.Value * 240);
                pPr.Append($"<w:spacing w:before=\"{before}\" w:after=\"{after}\" w:line=\"{line}\" w:lineRule=\"auto\"/>");
            }
            else
            {
                pPr.Append($"<w:spacing w:before=\"{before}\" w:after=\"{after}\"/>");
            }
        }

        if (text.LeftIndent.HasValue || text.RightIndent.HasValue || text.FirstLineIndent.HasValue || text.HangingIndent.HasValue)
        {
            pPr.Append("<w:ind");
            if (text.LeftIndent.HasValue)
                pPr.Append($" w:left=\"{ToDxa(text.LeftIndent.Value)}\"");
            if (text.RightIndent.HasValue)
                pPr.Append($" w:right=\"{ToDxa(text.RightIndent.Value)}\"");
            if (text.FirstLineIndent.HasValue)
                pPr.Append($" w:firstLine=\"{ToDxa(text.FirstLineIndent.Value)}\"");
            if (text.HangingIndent.HasValue)
                pPr.Append($" w:hanging=\"{ToDxa(text.HangingIndent.Value)}\"");
            pPr.Append("/>");
        }

        string alignment = text.Alignment != "left"
            ? text.Alignment
            : useDefaultTextStyle ? _defaultTextStyle?.Alignment ?? "left" : "left";

        string jc = alignment switch
        {
            "center" => "center",
            "right"  => "right",
            "both"   => "both",
            _        => "left"
        };
        if (jc != "left")
            pPr.Append($"<w:jc w:val=\"{jc}\"/>");

        if (pPr.Length > 0)
        {
            _sb.Append("<w:pPr>");
            _sb.Append(pPr);
            _sb.Append("</w:pPr>");
        }

        var bookmarkId = -1;
        if (!string.IsNullOrWhiteSpace(text.BookmarkName))
        {
            bookmarkId = _ctx.NextBookmarkId();
            _sb.Append($"<w:bookmarkStart w:id=\"{bookmarkId}\" w:name=\"{TextDescriptor.BookmarkName(text.BookmarkName)}\"/>");
        }

        var paragraphDefaultRun = text.Runs.Count > 0 ? text.Runs[0] : null;
        foreach (var run in text.Runs)
            WriteRun(run, paragraphDefaultRun, useDefaultTextStyle);

        if (bookmarkId >= 0)
            _sb.Append($"<w:bookmarkEnd w:id=\"{bookmarkId}\"/>");

        _sb.AppendLine("</w:p>");
    }

    private void WriteRun(TextRun run, TextRun? paragraphDefaultRun, bool useDefaultTextStyle = true)
    {
        if (run.Kind == TextRunKind.Text && run.Text.Length == 0)
            return;

        string rPr = BuildRPr(run, paragraphDefaultRun, useDefaultTextStyle);

        switch (run.Kind)
        {
            case TextRunKind.Tab:
                _sb.Append($"<w:r>{rPr}<w:tab/></w:r>");
                break;
            case TextRunKind.Hyperlink:
                WriteHyperlinkRun(run, rPr);
                break;
            case TextRunKind.Field:
                WriteFieldRun(rPr, run.FieldInstruction ?? string.Empty, run.FieldCachedText ?? string.Empty);
                break;
            case TextRunKind.Footnote:
                WriteFootnoteReference(run, rPr);
                break;
            case TextRunKind.Endnote:
                WriteEndnoteReference(run, rPr);
                break;
            case TextRunKind.PageNumber:
                WriteFieldRun(rPr, "PAGE", "1");
                break;
            case TextRunKind.PageCount:
                WriteFieldRun(rPr, "NUMPAGES", "1");
                break;
            default:
                _sb.Append("<w:r>");
                _sb.Append(rPr);
                _sb.Append($"<w:t xml:space=\"preserve\">{OoxmlWriter.Escape(run.Text)}</w:t>");
                _sb.Append("</w:r>");
                break;
        }
    }

    private void WriteHyperlinkRun(TextRun run, string rPr)
    {
        if (string.IsNullOrWhiteSpace(run.Url))
            return;

        var rId = _relationships.AddHyperlink(run.Url);
        _sb.Append($"<w:hyperlink r:id=\"{rId}\" w:history=\"1\"><w:r>");
        _sb.Append(rPr);
        _sb.Append($"<w:t xml:space=\"preserve\">{OoxmlWriter.Escape(run.Text)}</w:t>");
        _sb.Append("</w:r></w:hyperlink>");
    }

    private void WriteFootnoteReference(TextRun run, string rPr)
    {
        if (run.NoteText == null)
            return;

        var id = _ctx.AddFootnote(run.NoteText);
        _sb.Append($"<w:r>{rPr}<w:footnoteReference w:id=\"{id}\"/></w:r>");
    }

    private void WriteEndnoteReference(TextRun run, string rPr)
    {
        if (run.NoteText == null)
            return;

        var id = _ctx.AddEndnote(run.NoteText);
        _sb.Append($"<w:r>{rPr}<w:endnoteReference w:id=\"{id}\"/></w:r>");
    }

    private string BuildRPr(TextRun run, TextRun? paragraphDefaultRun, bool useDefaultTextStyle)
    {
        // Element order follows CT_RPr sequence: rFonts, b, i, strike, color, sz, szCs, highlight, u, vertAlign
        var sb = new StringBuilder();
        var defaultRun = useDefaultTextStyle ? _defaultRun : null;
        string? fontFamily = run.FontFamily ?? paragraphDefaultRun?.FontFamily ?? defaultRun?.FontFamily;
        bool bold = run.Bold ?? paragraphDefaultRun?.Bold ?? defaultRun?.Bold ?? false;
        bool italic = run.Italic ?? paragraphDefaultRun?.Italic ?? defaultRun?.Italic ?? false;
        bool strikethrough = run.Strikethrough ?? paragraphDefaultRun?.Strikethrough ?? defaultRun?.Strikethrough ?? false;
        bool underline = run.Underline ?? paragraphDefaultRun?.Underline ?? defaultRun?.Underline ?? false;
        string? fontColor = run.FontColor ?? paragraphDefaultRun?.FontColor ?? defaultRun?.FontColor;
        float? fontSize = run.FontSize ?? paragraphDefaultRun?.FontSize ?? defaultRun?.FontSize;
        string? highlightColor = run.HighlightColor ?? paragraphDefaultRun?.HighlightColor ?? defaultRun?.HighlightColor;
        string? verticalAlignment = run.VerticalAlignment ?? paragraphDefaultRun?.VerticalAlignment ?? defaultRun?.VerticalAlignment;

        if (fontFamily != null)
            sb.Append($"<w:rFonts w:ascii=\"{OoxmlWriter.Escape(fontFamily)}\" w:hAnsi=\"{OoxmlWriter.Escape(fontFamily)}\"/>");
        if (bold)          sb.Append("<w:b/>");
        if (italic)        sb.Append("<w:i/>");
        if (strikethrough) sb.Append("<w:strike/>");
        if (fontColor != null)
            sb.Append($"<w:color w:val=\"{fontColor.TrimStart('#')}\"/>");
        if (fontSize.HasValue)
        {
            int hp = (int)(fontSize.Value * 2);
            sb.Append($"<w:sz w:val=\"{hp}\"/><w:szCs w:val=\"{hp}\"/>");
        }
        if (highlightColor != null)
            sb.Append($"<w:highlight w:val=\"{highlightColor}\"/>");
        if (underline)     sb.Append("<w:u w:val=\"single\"/>");
        if (verticalAlignment != null)
            sb.Append($"<w:vertAlign w:val=\"{verticalAlignment}\"/>");
        return sb.Length == 0 ? string.Empty : $"<w:rPr>{sb}</w:rPr>";
    }

    // A OOXML field run: begin → instrText → separate → cached value → end
    private void WriteFieldRun(string rPr, string fieldInstruction, string cachedText)
    {
        _sb.Append($"<w:r>{rPr}<w:fldChar w:fldCharType=\"begin\"/></w:r>");
        _sb.Append($"<w:r>{rPr}<w:instrText xml:space=\"preserve\"> {OoxmlWriter.Escape(fieldInstruction)} </w:instrText></w:r>");
        _sb.Append($"<w:r>{rPr}<w:fldChar w:fldCharType=\"separate\"/></w:r>");
        _sb.Append($"<w:r>{rPr}<w:t xml:space=\"preserve\">{OoxmlWriter.Escape(cachedText)}</w:t></w:r>");
        _sb.Append($"<w:r>{rPr}<w:fldChar w:fldCharType=\"end\"/></w:r>");
    }

    // -------------------------------------------------------------------------
    // Layout
    // -------------------------------------------------------------------------

    private void WriteColumn(ColumnElement col)
    {
        for (int i = 0; i < col.Items.Count; i++)
        {
            WriteContainer(col.Items[i]);
            if (col.Spacing > 0 && i < col.Items.Count - 1)
                WriteSpacingParagraph(col.Spacing);
        }
    }

    private void WriteSpacingParagraph(float points)
    {
        int after = (int)(points * 20);
        _sb.AppendLine($"<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"{after}\"/></w:pPr></w:p>");
    }

    private void WriteList(ListElement list)
    {
        int numId = _ctx.AddList(list);
        foreach (var item in list.Items)
            WriteParagraph(item.Text, numId, item.Level);
    }

    private void WriteBookmark(BookmarkElement bookmark)
    {
        var bookmarkId = _ctx.NextBookmarkId();
        var name = TextDescriptor.BookmarkName(bookmark.Name);
        _sb.AppendLine($"<w:p><w:bookmarkStart w:id=\"{bookmarkId}\" w:name=\"{name}\"/><w:bookmarkEnd w:id=\"{bookmarkId}\"/></w:p>");
    }

    private void WriteTableOfContents(TableOfContentsElement toc)
    {
        if (!string.IsNullOrWhiteSpace(toc.Title))
        {
            var title = new TextElement { StyleId = "Heading1" };
            title.Runs.Add(new TextRun { Text = toc.Title });
            WriteParagraph(title);
        }

        var min = MathCompat.Clamp(toc.MinLevel, 1, 9);
        var max = MathCompat.Clamp(toc.MaxLevel, min, 9);
        _sb.Append("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>");
        _sb.Append($"<w:r><w:instrText xml:space=\"preserve\"> TOC \\o &quot;{min}-{max}&quot; \\h \\z \\u </w:instrText></w:r>");
        _sb.Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
        _sb.Append("<w:r><w:t>Right-click to update field.</w:t></w:r>");
        _sb.AppendLine("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
    }

    private void WriteRow(RowElement row)
    {
        float totalRelative = row.Cells
            .Where(c => c.Mode == RowCell.SizingMode.Relative)
            .Sum(c => c.Size);
        if (totalRelative == 0) totalRelative = 1;

        // Reserve pct budget for auto cells so relative cells don't claim the full width.
        // Each auto cell reserves 20% (1000 of 5000), leaving the rest for relative cells.
        int numAutoCells = row.Cells.Count(c => c.Mode == RowCell.SizingMode.Auto);
        int pctForRelative = Math.Max(1000, 5000 - numAutoCells * 1000);
        int cellSpacing = Math.Max(0, (int)(row.Spacing * 20));

        _sb.AppendLine("<w:tbl>");
        _sb.Append("<w:tblPr>");
        _sb.Append("<w:tblW w:w=\"5000\" w:type=\"pct\"/>");
        if (cellSpacing > 0)
            _sb.Append($"<w:tblCellSpacing w:w=\"{cellSpacing}\" w:type=\"dxa\"/>");
        _sb.AppendLine("</w:tblPr>");
        _sb.Append("<w:tblGrid>");
        foreach (var cell in row.Cells) _sb.Append($"<w:gridCol w:w=\"{RowGridWidth(cell.Mode, cell.Size, totalRelative, pctForRelative)}\"/>");
        _sb.AppendLine("</w:tblGrid>");

        _sb.AppendLine("<w:tr>");
        foreach (var cell in row.Cells)
        {
            string widthAttr = cell.Mode switch
            {
                RowCell.SizingMode.Constant => $"w:w=\"{(int)(cell.Size * 20)}\" w:type=\"dxa\"",
                RowCell.SizingMode.Auto     => "w:w=\"0\" w:type=\"auto\"",
                _                           => $"w:w=\"{(int)(cell.Size / totalRelative * pctForRelative)}\" w:type=\"pct\""
            };
            _sb.Append($"<w:tc><w:tcPr><w:tcW {widthAttr}/></w:tcPr>");
            WriteContainer(cell.Container);
            EnsureCellHasParagraph(cell.Container);
            _sb.AppendLine("</w:tc>");
        }
        _sb.AppendLine("</w:tr></w:tbl>");
        _sb.AppendLine("<w:p/>");
    }

    private void WriteTable(TableElement table)
    {
        EnsureTableColumns(table);

        float totalRelative = table.Columns
            .Where(c => c.Mode == TableColumnDef.SizingMode.Relative)
            .Sum(c => c.Size);
        if (totalRelative == 0) totalRelative = 1;

        var mixed = GetMixedColumnSizing(table);

        _sb.AppendLine("<w:tbl>");
        _sb.Append("<w:tblPr>");
        if (!string.IsNullOrWhiteSpace(table.StyleId))
            _sb.Append($"<w:tblStyle w:val=\"{table.StyleId}\"/>");
        _sb.Append(TableWidthXml(table));
        if (table.Alignment != "left")
            _sb.Append($"<w:jc w:val=\"{table.Alignment}\"/>");
        if (table.BorderWidth > 0)
        {
            string bw = ((int)(table.BorderWidth * 8)).ToString();
            string bc = table.BorderColor.TrimStart('#');
            _sb.Append($"<w:tblBorders>" +
                       $"<w:top w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"<w:left w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"<w:bottom w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"<w:right w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"<w:insideH w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"<w:insideV w:val=\"single\" w:sz=\"{bw}\" w:space=\"0\" w:color=\"{bc}\"/>" +
                       $"</w:tblBorders>");
        }
        _sb.Append($"<w:tblCellMar>" +
                   $"<w:top w:w=\"{ToDxa(table.CellPaddingTop)}\" w:type=\"dxa\"/>" +
                   $"<w:left w:w=\"{ToDxa(table.CellPaddingLeft)}\" w:type=\"dxa\"/>" +
                   $"<w:bottom w:w=\"{ToDxa(table.CellPaddingBottom)}\" w:type=\"dxa\"/>" +
                   $"<w:right w:w=\"{ToDxa(table.CellPaddingRight)}\" w:type=\"dxa\"/>" +
                   $"</w:tblCellMar>");
        _sb.Append("<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/>");
        _sb.AppendLine("</w:tblPr>");

        _sb.Append("<w:tblGrid>");
        foreach (var col in table.Columns) _sb.Append($"<w:gridCol w:w=\"{GridWidth(col, totalRelative, mixed)}\"/>");
        _sb.AppendLine("</w:tblGrid>");

        foreach (var trow in table.HeaderRows)
            WriteTableRow(table, trow, totalRelative, mixed, isHeader: true, useAlternateBackground: false);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            WriteTableRow(table, table.Rows[rowIndex], totalRelative, mixed, isHeader: false, useAlternateBackground: rowIndex % 2 == 1);

        _sb.AppendLine("</w:tbl><w:p/>");
    }

    private void WriteTableRow(TableElement table, TableRow trow, float totalRelative, MixedColumnSizing? mixed, bool isHeader, bool useAlternateBackground)
    {
        _sb.AppendLine("<w:tr>");
        var rowMinHeight = isHeader && table.HeaderRowMinHeight > 0
            ? table.HeaderRowMinHeight
            : table.RowMinHeight;

        if (trow.KeepTogether || isHeader || rowMinHeight > 0)
        {
            _sb.Append("<w:trPr>");
            if (trow.KeepTogether)
                _sb.Append("<w:cantSplit/>");
            if (rowMinHeight > 0)
                _sb.Append($"<w:trHeight w:val=\"{ToDxa(rowMinHeight)}\" w:hRule=\"atLeast\"/>");
            if (isHeader)
                _sb.Append("<w:tblHeader/>");
            _sb.Append("</w:trPr>");
        }

        int columnIndex = 0;
        foreach (var cell in trow.Cells)
        {
            if (columnIndex >= table.Columns.Count)
                break;

            int span = Math.Min(Math.Max(1, cell.ColumnSpan), table.Columns.Count - columnIndex);
            string widthAttr = TableCellWidth(table, columnIndex, span, totalRelative, mixed);

            var fill = isHeader
                ? table.HeaderBackgroundColor
                : useAlternateBackground ? table.AlternateRowBackgroundColor : null;
            fill = cell.BackgroundColor ?? fill;

            _sb.Append($"<w:tc><w:tcPr><w:tcW {widthAttr}/>");
            if (span > 1)
                _sb.Append($"<w:gridSpan w:val=\"{span}\"/>");
            if (cell.VerticalMerge != null)
                _sb.Append(cell.VerticalMerge == "restart" ? "<w:vMerge w:val=\"restart\"/>" : "<w:vMerge/>");
            WriteCellBorders(cell);
            if (!string.IsNullOrWhiteSpace(fill))
                _sb.Append($"<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{fill.TrimStart('#')}\"/>");
            if (HasCellPaddingOverride(cell))
            {
                _sb.Append("<w:tcMar>");
                if (cell.PaddingTop.HasValue)
                    _sb.Append($"<w:top w:w=\"{ToDxa(cell.PaddingTop.Value)}\" w:type=\"dxa\"/>");
                if (cell.PaddingLeft.HasValue)
                    _sb.Append($"<w:left w:w=\"{ToDxa(cell.PaddingLeft.Value)}\" w:type=\"dxa\"/>");
                if (cell.PaddingBottom.HasValue)
                    _sb.Append($"<w:bottom w:w=\"{ToDxa(cell.PaddingBottom.Value)}\" w:type=\"dxa\"/>");
                if (cell.PaddingRight.HasValue)
                    _sb.Append($"<w:right w:w=\"{ToDxa(cell.PaddingRight.Value)}\" w:type=\"dxa\"/>");
                _sb.Append("</w:tcMar>");
            }
            if (!string.IsNullOrWhiteSpace(cell.TextDirection))
                _sb.Append($"<w:textDirection w:val=\"{cell.TextDirection}\"/>");
            if (!string.IsNullOrWhiteSpace(cell.VerticalAlignment))
                _sb.Append($"<w:vAlign w:val=\"{cell.VerticalAlignment}\"/>");
            _sb.Append("</w:tcPr>");
            WriteContainer(cell.Container);
            EnsureCellHasParagraph(cell.Container);
            _sb.AppendLine("</w:tc>");
            columnIndex += span;
        }
        _sb.AppendLine("</w:tr>");
    }

    // If ColumnsDefinition() was never called, infer relative columns from the widest row.
    // Otherwise, fail fast if any row's cells (accounting for column spans) exceed the
    // defined column count instead of silently dropping the overflowing cells.
    private static void EnsureTableColumns(TableElement table)
    {
        if (table.Columns.Count > 0)
        {
            ValidateRowSpans(table);
            return;
        }

        var maxSpan = table.HeaderRows.Concat(table.Rows).Select(RowSpan).DefaultIfEmpty(0).Max();
        for (int i = 0; i < maxSpan; i++)
            table.Columns.Add(new TableColumnDef { Mode = TableColumnDef.SizingMode.Relative, Size = 1 });
    }

    private static void ValidateRowSpans(TableElement table)
    {
        for (int i = 0; i < table.HeaderRows.Count; i++)
            ValidateRowSpan(table, table.HeaderRows[i], $"header row {i + 1}");

        for (int i = 0; i < table.Rows.Count; i++)
            ValidateRowSpan(table, table.Rows[i], $"row {i + 1}");
    }

    private static void ValidateRowSpan(TableElement table, TableRow trow, string rowLabel)
    {
        var span = RowSpan(trow);
        if (span > table.Columns.Count)
            throw new InvalidOperationException(
                $"Table {rowLabel} has {trow.Cells.Count} cell(s) spanning {span} column(s), " +
                $"but the table only defines {table.Columns.Count} column(s) via ColumnsDefinition(). " +
                $"Add more columns or reduce the cells/column spans in this row.");
    }

    private static int RowSpan(TableRow row) =>
        row.Cells.Sum(c => Math.Max(1, c.ColumnSpan));

    private void WriteChart(ChartElement chart)
    {
        var rId = _relationships.AddChart(chart);
        var docPrId = _ctx.NextDocPrId();
        const long cx = 5486400; // 6 inches
        const long cy = 3200400; // 3.5 inches

        var chartKind = chart.Series.FirstOrDefault()?.Kind ?? "bar";
        _sb.Append("<w:p><w:r><w:drawing>");
        _sb.Append("<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        _sb.Append($"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>");
        _sb.Append("<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>");
        _sb.Append($"<wp:docPr id=\"{docPrId}\" name=\"{OoxmlWriter.Escape(chart.Title)}\" descr=\"{OoxmlWriter.Escape(chartKind)} chart\"/>");
        _sb.Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
        _sb.Append("<a:graphic>");
        _sb.Append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">");
        _sb.Append($"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"{rId}\"/>");
        _sb.Append("</a:graphicData>");
        _sb.Append("</a:graphic>");
        _sb.Append("</wp:inline>");
        _sb.AppendLine("</w:drawing></w:r></w:p>");
    }

    // -------------------------------------------------------------------------
    // Image
    // -------------------------------------------------------------------------

    private void WriteImage(ImageElement img)
    {
        if (!ImageExists(img))
        {
            _sb.AppendLine($"<w:p><w:r><w:t>[Missing image: {OoxmlWriter.Escape(img.FilePath ?? img.FileName)}]</w:t></w:r></w:p>");
            return;
        }

        string rId = _relationships.AddImage(img);
        int docPrId = _ctx.NextDocPrId();
        string name = $"Image{docPrId}";

        var (cx, cy) = ImageExtents(img);
        var isFloating = img.WrapMode != "inline";

        _sb.Append("<w:p>");
        var jc = isFloating ? null : ImageJustification(img.Alignment);
        if (jc != null)
            _sb.Append($"<w:pPr><w:jc w:val=\"{jc}\"/></w:pPr>");
        _sb.Append("<w:r><w:drawing>");
        if (isFloating)
            WriteFloatingImage(img, rId, docPrId, name, cx, cy);
        else
            WriteInlineImage(img, rId, docPrId, name, cx, cy);
        _sb.AppendLine("</w:drawing></w:r></w:p>");

        if (img.Caption != null)
            WriteParagraph(img.Caption);
    }

    private void WriteInlineImage(ImageElement img, string rId, int docPrId, string name, long cx, long cy)
    {
        _sb.Append($"<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        _sb.Append($"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>");
        _sb.Append("<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>");
        WriteDrawingProperties(img, docPrId, name);
        _sb.Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
        WritePictureGraphic(img, rId, docPrId, name, cx, cy);
        _sb.Append("</wp:inline>");
    }

    private void WriteFloatingImage(ImageElement img, string rId, int docPrId, string name, long cx, long cy)
    {
        _sb.Append($"<wp:anchor distT=\"{ToEmu(img.MarginTop)}\" distB=\"{ToEmu(img.MarginBottom)}\" distL=\"{ToEmu(img.MarginLeft)}\" distR=\"{ToEmu(img.MarginRight)}\" simplePos=\"0\" relativeHeight=\"251659264\" behindDoc=\"{(img.WrapMode == "behind" ? "1" : "0")}\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">");
        _sb.Append("<wp:simplePos x=\"0\" y=\"0\"/>");
        WriteFloatingPosition("H", img.HorizontalRelativeFrom, FloatingHorizontalAlign(img), img.HorizontalOffset);
        WriteFloatingPosition("V", img.VerticalRelativeFrom, img.VerticalAlignment, img.VerticalOffset);
        _sb.Append($"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>");
        _sb.Append("<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>");
        WriteWrapMode(img);
        WriteDrawingProperties(img, docPrId, name);
        _sb.Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
        WritePictureGraphic(img, rId, docPrId, name, cx, cy);
        _sb.Append("</wp:anchor>");
    }

    private void WriteDrawingProperties(ImageElement img, int docPrId, string name)
    {
        _sb.Append($"<wp:docPr id=\"{docPrId}\" name=\"{OoxmlWriter.Escape(name)}\"");
        if (!string.IsNullOrWhiteSpace(img.AltText))
            _sb.Append($" descr=\"{OoxmlWriter.Escape(img.AltText)}\"");
        _sb.Append("/>");
    }

    // When neither Position()/PositionFromPage() (HorizontalOffset) nor a Float*() alignment
    // (HorizontalAlignment) was set, fall back to the image's Alignment() so AlignCenter()/
    // AlignRight() are honored for floating images too, instead of always sitting at x=0.
    private static string? FloatingHorizontalAlign(ImageElement img)
    {
        if (img.HorizontalAlignment != null || img.HorizontalOffset.HasValue)
            return img.HorizontalAlignment;

        return img.Alignment ?? DefaultFloatingAlignment(img.WrapMode);
    }

    // Top/bottom-wrapped and behind/in-front images are typically full-width banners or
    // background art, so center them by default; square/tight wraps keep flush-left since
    // text wraps around them on whichever side they land.
    private static string DefaultFloatingAlignment(string wrapMode) =>
        wrapMode is "topBottom" or "behind" or "inFront" ? "center" : "left";

    private void WriteFloatingPosition(string axis, string relativeFrom, string? alignment, float? offset)
    {
        _sb.Append($"<wp:position{axis} relativeFrom=\"{relativeFrom}\">");
        if (!string.IsNullOrWhiteSpace(alignment))
            _sb.Append($"<wp:align>{alignment}</wp:align>");
        else
            _sb.Append($"<wp:posOffset>{ToEmu(offset ?? 0)}</wp:posOffset>");
        _sb.Append($"</wp:position{axis}>");
    }

    private void WriteWrapMode(ImageElement img)
    {
        switch (img.WrapMode)
        {
            case "tight":
                _sb.Append("<wp:wrapTight wrapText=\"bothSides\"><wp:wrapPolygon edited=\"0\"><wp:start x=\"0\" y=\"0\"/><wp:lineTo x=\"0\" y=\"21600\"/><wp:lineTo x=\"21600\" y=\"21600\"/><wp:lineTo x=\"21600\" y=\"0\"/><wp:lineTo x=\"0\" y=\"0\"/></wp:wrapPolygon></wp:wrapTight>");
                break;
            case "topBottom":
                _sb.Append("<wp:wrapTopAndBottom/>");
                break;
            case "behind":
            case "inFront":
                _sb.Append("<wp:wrapNone/>");
                break;
            default:
                _sb.Append("<wp:wrapSquare wrapText=\"bothSides\"/>");
                break;
        }
    }

    private void WritePictureGraphic(ImageElement img, string rId, int docPrId, string name, long cx, long cy)
    {
        _sb.Append("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        _sb.Append("<pic:pic>");
        _sb.Append($"<pic:nvPicPr><pic:cNvPr id=\"{docPrId}\" name=\"{OoxmlWriter.Escape(name)}\"");
        if (!string.IsNullOrWhiteSpace(img.AltText))
            _sb.Append($" descr=\"{OoxmlWriter.Escape(img.AltText)}\"");
        _sb.Append("/><pic:cNvPicPr/></pic:nvPicPr>");
        _sb.Append($"<pic:blipFill><a:blip r:embed=\"{rId}\"/>");
        WriteCrop(img);
        _sb.Append("<a:stretch><a:fillRect/></a:stretch></pic:blipFill>");
        _sb.Append($"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>");
        _sb.Append($"<a:prstGeom prst=\"{(img.Rounded ? "roundRect" : "rect")}\"><a:avLst/></a:prstGeom>");
        WriteImageBorder(img);
        _sb.Append("</pic:spPr>");
        _sb.Append("</pic:pic>");
        _sb.Append("</a:graphicData></a:graphic>");
    }

    private void WriteCrop(ImageElement img)
    {
        if (img.CropLeftPercent <= 0 && img.CropTopPercent <= 0 && img.CropRightPercent <= 0 && img.CropBottomPercent <= 0)
            return;

        _sb.Append("<a:srcRect");
        if (img.CropLeftPercent > 0)
            _sb.Append($" l=\"{ToDrawingPercent(img.CropLeftPercent)}\"");
        if (img.CropTopPercent > 0)
            _sb.Append($" t=\"{ToDrawingPercent(img.CropTopPercent)}\"");
        if (img.CropRightPercent > 0)
            _sb.Append($" r=\"{ToDrawingPercent(img.CropRightPercent)}\"");
        if (img.CropBottomPercent > 0)
            _sb.Append($" b=\"{ToDrawingPercent(img.CropBottomPercent)}\"");
        _sb.Append("/>");
    }

    private void WriteImageBorder(ImageElement img)
    {
        if (!img.BorderWidth.HasValue || img.BorderWidth.Value <= 0)
            return;

        _sb.Append($"<a:ln w=\"{ToEmu(img.BorderWidth.Value)}\"><a:solidFill><a:srgbClr val=\"{img.BorderColor.TrimStart('#')}\"/></a:solidFill></a:ln>");
    }

    private static bool ImageExists(ImageElement img) =>
        img.Bytes is { Length: > 0 } ||
        (!string.IsNullOrWhiteSpace(img.FilePath) && File.Exists(img.FilePath));

    private static (long cx, long cy) ImageExtents(ImageElement img)
    {
        var (px, py) = img.Bytes != null
            ? ImageReader.GetPixelDimensions(img.Bytes, img.FileName)
            : ImageReader.GetPixelDimensions(img.FilePath!);

        double aspect = px > 0 && py > 0 ? (double)py / px : 1d;
        double widthPoints;
        double heightPoints;

        if (img.Width.HasValue && img.Height.HasValue)
        {
            widthPoints = img.Width.Value;
            heightPoints = img.Height.Value;
        }
        else if (img.Width.HasValue)
        {
            widthPoints = img.Width.Value;
            heightPoints = widthPoints * aspect;
        }
        else if (img.Height.HasValue)
        {
            heightPoints = img.Height.Value;
            widthPoints = aspect > 0 ? heightPoints / aspect : heightPoints;
        }
        else if (px > 0 && py > 0)
        {
            widthPoints = px * 72d / 96d;
            heightPoints = py * 72d / 96d;
        }
        else
        {
            widthPoints = 90;
            heightPoints = 67.5;
        }

        if (img.MaxWidth.HasValue && widthPoints > img.MaxWidth.Value)
        {
            var scale = img.MaxWidth.Value / widthPoints;
            widthPoints *= scale;
            heightPoints *= scale;
        }

        return ((long)(widthPoints * 12700), (long)(heightPoints * 12700));
    }

    private static string? ImageJustification(string? alignment) =>
        alignment switch
        {
            "center" => "center",
            "right" => "right",
            _ => null
        };

    private static void WriteParagraphBorders(TextElement text, StringBuilder pPr)
    {
        if (text.TopBorder == null && text.RightBorder == null && text.BottomBorder == null && text.LeftBorder == null)
            return;

        pPr.Append("<w:pBdr>");
        WriteParagraphBorder("top", text.TopBorder, pPr);
        WriteParagraphBorder("left", text.LeftBorder, pPr);
        WriteParagraphBorder("bottom", text.BottomBorder, pPr);
        WriteParagraphBorder("right", text.RightBorder, pPr);
        pPr.Append("</w:pBdr>");
    }

    private static void WriteParagraphBorder(string side, ParagraphBorder? border, StringBuilder pPr)
    {
        if (border == null)
            return;

        var size = Math.Max(1, (int)(border.Width * 8));
        var space = Math.Max(0, (int)border.Space);
        pPr.Append($"<w:{side} w:val=\"single\" w:sz=\"{size}\" w:space=\"{space}\" w:color=\"{border.Color.TrimStart('#')}\"/>");
    }

    private static long ToEmu(float points) => Math.Max(0, (long)(points * 12700));

    private static int ToDrawingPercent(float percent) => MathCompat.Clamp((int)(percent * 1000), 0, 100000);

    // -------------------------------------------------------------------------
    // Misc
    // -------------------------------------------------------------------------

    private void WriteHorizontalLine()
    {
        _sb.AppendLine("""
            <w:p>
              <w:pPr>
                <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>
              </w:pPr>
            </w:p>
            """);
    }

    private void WritePageBreak()
    {
        _sb.AppendLine("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>");
    }

    private static int RowGridWidth(RowCell.SizingMode mode, float size, float totalRelative, int pctForRelative) =>
        mode switch
        {
            RowCell.SizingMode.Constant => Math.Max(1, (int)(size * 20)),
            RowCell.SizingMode.Auto     => 5000 - pctForRelative,
            _                           => Math.Max(1, (int)(size / totalRelative * pctForRelative))
        };

    private static int GridWidth(TableColumnDef col, float totalRelative, MixedColumnSizing? mixed)
    {
        if (mixed is { } m)
            return Math.Max(1, ToDxa(ColumnPoints(col, totalRelative, m.RemainingPoints)));

        return col.Mode == TableColumnDef.SizingMode.Constant
            ? Math.Max(1, (int)(col.Size * 20))
            : Math.Max(1, (int)(col.Size / totalRelative * 5000));
    }

    // Assumed page content width (Letter, 8.5in page minus 1in margins on each side) used as the
    // "100%" reference for mixed Relative/Constant column tables when no absolute Width(points)
    // was set, so gridCol/tcW can be expressed in consistent units instead of mixing dxa and pct.
    private const float DefaultContentWidthPoints = 468f;

    private static bool HasMixedColumnSizing(TableElement table) =>
        table.Columns.Any(c => c.Mode == TableColumnDef.SizingMode.Relative) &&
        table.Columns.Any(c => c.Mode == TableColumnDef.SizingMode.Constant);

    // RemainingPoints is the width left for relative columns after constant columns claim their
    // exact widths; TotalPoints (= sum of constant widths + RemainingPoints) is the shared
    // percentage denominator for both gridCol (as dxa) and tcW (as pct) in a mixed table.
    private readonly record struct MixedColumnSizing(float RemainingPoints, float TotalPoints);

    private static MixedColumnSizing? GetMixedColumnSizing(TableElement table)
    {
        if (!HasMixedColumnSizing(table))
            return null;

        float constantTotal = table.Columns
            .Where(c => c.Mode == TableColumnDef.SizingMode.Constant)
            .Sum(c => c.Size);
        float reference = table.Width.HasValue && table.WidthType == "dxa" ? table.Width.Value : DefaultContentWidthPoints;
        float remaining = Math.Max(1f, reference - constantTotal);
        return new MixedColumnSizing(remaining, constantTotal + remaining);
    }

    private static float ColumnPoints(TableColumnDef col, float totalRelative, float remainingPoints) =>
        col.Mode == TableColumnDef.SizingMode.Constant
            ? col.Size
            : col.Size / totalRelative * remainingPoints;

    private static int ToDxa(float points) => Math.Max(0, (int)(points * 20));

    private static string TabStopAlignment(TabStopAlignment alignment) =>
        alignment switch
        {
            TerraFluent.Docx.Reporting.TabStopAlignment.Center => "center",
            TerraFluent.Docx.Reporting.TabStopAlignment.Right => "right",
            TerraFluent.Docx.Reporting.TabStopAlignment.Decimal => "decimal",
            _ => "left"
        };

    private static string TableWidthXml(TableElement table)
    {
        if (!table.Width.HasValue)
            return "<w:tblW w:w=\"5000\" w:type=\"pct\"/>";

        return table.WidthType == "dxa"
            ? $"<w:tblW w:w=\"{ToDxa(table.Width.Value)}\" w:type=\"dxa\"/>"
            : $"<w:tblW w:w=\"{Math.Max(1, (int)(table.Width.Value * 50))}\" w:type=\"pct\"/>";
    }

    private void WriteCellBorders(TableCell cell)
    {
        if (cell.TopBorder == null && cell.RightBorder == null && cell.BottomBorder == null && cell.LeftBorder == null)
            return;

        _sb.Append("<w:tcBorders>");
        WriteCellBorder("top", cell.TopBorder);
        WriteCellBorder("left", cell.LeftBorder);
        WriteCellBorder("bottom", cell.BottomBorder);
        WriteCellBorder("right", cell.RightBorder);
        _sb.Append("</w:tcBorders>");
    }

    private void WriteCellBorder(string side, CellBorder? border)
    {
        if (border == null)
            return;

        var sz = Math.Max(1, (int)(border.Width * 8));
        _sb.Append($"<w:{side} w:val=\"single\" w:sz=\"{sz}\" w:space=\"0\" w:color=\"{border.Color.TrimStart('#')}\"/>");
    }

    private static bool HasCellPaddingOverride(TableCell cell) =>
        cell.PaddingTop.HasValue ||
        cell.PaddingRight.HasValue ||
        cell.PaddingBottom.HasValue ||
        cell.PaddingLeft.HasValue;

    private static string TableCellWidth(TableElement table, int startColumn, int span, float totalRelative, MixedColumnSizing? mixed)
    {
        var columns = table.Columns.Skip(startColumn).Take(span).ToList();

        if (mixed is { } m)
        {
            var points = columns.Sum(c => ColumnPoints(c, totalRelative, m.RemainingPoints));
            return $"w:w=\"{Math.Max(1, (int)(points / m.TotalPoints * 5000))}\" w:type=\"pct\"";
        }

        if (columns.All(c => c.Mode == TableColumnDef.SizingMode.Constant))
            return $"w:w=\"{ToDxa(columns.Sum(c => c.Size))}\" w:type=\"dxa\"";

        var pct = columns.Sum(c => c.Mode == TableColumnDef.SizingMode.Relative
            ? c.Size / totalRelative * 5000
            : c.Size * 20);
        return $"w:w=\"{Math.Max(1, (int)pct)}\" w:type=\"pct\"";
    }

    private void EnsureCellHasParagraph(ContainerElement container)
    {
        if (container.Elements.Count == 0)
            _sb.Append("<w:p/>");
    }
}
