using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using TerraFluent.Docx.Reporting.Core;

namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Opens an existing DOCX file and replaces simple placeholders or tagged content controls.
/// </summary>
public sealed class DocxTemplate
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Xml = XNamespace.Xml;

    private readonly string _templatePath;
    private readonly Dictionary<string, string> _placeholders = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _contentControls = new(StringComparer.OrdinalIgnoreCase);

    private DocxTemplate(string templatePath)
    {
        _templatePath = templatePath;
    }

    /// <summary>
    /// Opens a DOCX template for replacement.
    /// </summary>
    /// <param name="templatePath">Path to an existing `.docx` template.</param>
    /// <returns>A template replacement builder.</returns>
    public static DocxTemplate Open(string templatePath)
    {
        Guard.NotWhiteSpace(templatePath, nameof(templatePath));
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("DOCX template file could not be found.", templatePath);

        return new DocxTemplate(templatePath);
    }

    /// <summary>
    /// Replaces text placeholders such as `{{Name}}` with a value.
    /// </summary>
    /// <param name="placeholder">Placeholder name with or without braces.</param>
    /// <param name="value">Replacement text.</param>
    /// <returns>The current template builder.</returns>
    public DocxTemplate Replace(string placeholder, string value)
    {
        _placeholders[PlaceholderToken(placeholder)] = Guard.NotNull(value, nameof(value));
        return this;
    }

    /// <summary>
    /// Replaces a content control by tag or alias.
    /// </summary>
    /// <param name="tagOrAlias">The content control tag or alias.</param>
    /// <param name="value">Replacement text.</param>
    /// <returns>The current template builder.</returns>
    public DocxTemplate ReplaceContentControl(string tagOrAlias, string value)
    {
        _contentControls[Guard.NotWhiteSpace(tagOrAlias, nameof(tagOrAlias))] = Guard.NotNull(value, nameof(value));
        return this;
    }

    /// <summary>
    /// Saves the replaced template to a new DOCX file.
    /// </summary>
    /// <param name="outputPath">Destination `.docx` path.</param>
    public void SaveAs(string outputPath)
    {
        Guard.NotWhiteSpace(outputPath, nameof(outputPath));

        using var source = ZipFile.OpenRead(_templatePath);
        using var output = File.Create(outputPath);
        using var dest = new ZipArchive(output, ZipArchiveMode.Create);

        foreach (var entry in source.Entries)
        {
            var newEntry = dest.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var entryStream = entry.Open();
            using var newEntryStream = newEntry.Open();

            if (IsEditableXml(entry.FullName))
            {
                using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                using var writer = new StreamWriter(newEntryStream, new UTF8Encoding(false));
                writer.Write(ApplyReplacements(reader.ReadToEnd()));
            }
            else
            {
                entryStream.CopyTo(newEntryStream);
            }
        }
    }

    /// <summary>
    /// Saves the replaced template to a DOCX byte array.
    /// </summary>
    /// <returns>The generated DOCX package bytes.</returns>
    public byte[] Save()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"terrafluent-docx-reporting-template-{Guid.NewGuid():N}.docx");
        try
        {
            SaveAs(tempPath);
            return File.ReadAllBytes(tempPath);
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    private string ApplyReplacements(string xml)
    {
        if (_placeholders.Count == 0 && _contentControls.Count == 0)
            return xml;

        var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);

        foreach (var contentControl in _contentControls)
            ReplaceContentControls(document, contentControl.Key, contentControl.Value);

        if (_placeholders.Count > 0)
        {
            foreach (var paragraph in document.Descendants(W + "p"))
                ReplacePlaceholders(paragraph, _placeholders);
        }

        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static void ReplaceContentControls(XDocument document, string key, string value)
    {
        foreach (var sdt in document.Descendants(W + "sdt").Where(s => ContentControlMatches(s, key)).ToList())
            ReplaceTextScope(sdt, value);
    }

    private static bool ContentControlMatches(XElement sdt, string key)
    {
        var properties = sdt.Element(W + "sdtPr");
        if (properties == null)
            return false;

        return properties.Elements(W + "tag")
            .Concat(properties.Elements(W + "alias"))
            .Any(e => string.Equals((string?)e.Attribute(W + "val"), key, StringComparison.OrdinalIgnoreCase));
    }

    private static void ReplaceTextScope(XElement scope, string value)
    {
        var textNodes = scope.Descendants(W + "t").ToList();
        if (textNodes.Count == 0)
            return;

        textNodes[0].Value = value;
        textNodes[0].SetAttributeValue(Xml + "space", "preserve");
        foreach (var textNode in textNodes.Skip(1))
            textNode.Value = string.Empty;
    }

    private static void ReplacePlaceholders(XElement scope, IReadOnlyDictionary<string, string> replacements)
    {
        var textNodes = scope.Descendants(W + "t").ToList();
        if (textNodes.Count == 0)
            return;

        var text = string.Concat(textNodes.Select(t => t.Value));
        foreach (var replacement in replacements)
        {
            var searchStart = 0;
            while (searchStart <= text.Length)
            {
                var index = text.IndexOf(replacement.Key, searchStart, StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                    break;

                ReplaceTextRange(textNodes, index, replacement.Key.Length, replacement.Value);
                text = string.Concat(textNodes.Select(t => t.Value));
                searchStart = index + replacement.Value.Length;
            }
        }
    }

    private static void ReplaceTextRange(IReadOnlyList<XElement> textNodes, int start, int length, string replacement)
    {
        var end = start + length;
        var position = 0;
        var replacementWritten = false;

        foreach (var node in textNodes)
        {
            var nodeText = node.Value;
            var nodeStart = position;
            var nodeEnd = position + nodeText.Length;
            position = nodeEnd;

            if (nodeEnd <= start || nodeStart >= end)
                continue;

            var prefixLength = Math.Max(0, start - nodeStart);
            var suffixStart = Math.Min(nodeText.Length, Math.Max(0, end - nodeStart));
            var prefix = prefixLength > 0 ? nodeText.Substring(0, prefixLength) : string.Empty;
            var suffix = suffixStart < nodeText.Length ? nodeText.Substring(suffixStart) : string.Empty;

            if (!replacementWritten)
            {
                node.Value = prefix + replacement + suffix;
                node.SetAttributeValue(Xml + "space", "preserve");
                replacementWritten = true;
            }
            else
            {
                node.Value = suffix;
            }
        }
    }

    private static string PlaceholderToken(string placeholder)
    {
        var normalized = Guard.NotWhiteSpace(placeholder, nameof(placeholder));
        return normalized.StartsWith("{{", StringComparison.Ordinal) && normalized.EndsWith("}}", StringComparison.Ordinal)
            ? normalized
            : "{{" + normalized + "}}";
    }

    private static bool IsEditableXml(string path) =>
        path.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
        (path.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) ||
        (path.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
}
