using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace TerraFluent.Docx.Reporting;

/// <summary>
/// Opens an existing DOCX file and replaces simple placeholders or tagged content controls.
/// </summary>
public sealed class DocxTemplate
{
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
    public static DocxTemplate Open(string templatePath) => new(templatePath);

    /// <summary>
    /// Replaces text placeholders such as `{{Name}}` with a value.
    /// </summary>
    /// <param name="placeholder">Placeholder name with or without braces.</param>
    /// <param name="value">Replacement text.</param>
    /// <returns>The current template builder.</returns>
    public DocxTemplate Replace(string placeholder, string value)
    {
        _placeholders[placeholder] = value;
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
        _contentControls[tagOrAlias] = value;
        return this;
    }

    /// <summary>
    /// Saves the replaced template to a new DOCX file.
    /// </summary>
    /// <param name="outputPath">Destination `.docx` path.</param>
    public void SaveAs(string outputPath)
    {
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
        foreach (var placeholder in _placeholders)
        {
            var escaped = Escape(placeholder.Value);
            xml = ReplaceIgnoreCase(xml, "{{" + placeholder.Key + "}}", escaped);
            xml = ReplaceIgnoreCase(xml, placeholder.Key, escaped);
        }

        if (_contentControls.Count == 0 || xml.IndexOf("<w:sdt", StringComparison.Ordinal) < 0)
            return xml;

        foreach (var contentControl in _contentControls)
            xml = ReplaceContentControl(xml, contentControl.Key, Escape(contentControl.Value));

        return xml;
    }

    // string.Replace(string, string, StringComparison) isn't part of netstandard2.0.
    private static string ReplaceIgnoreCase(string text, string oldValue, string newValue)
    {
        if (oldValue.Length == 0) return text;

        var sb = new StringBuilder();
        var index = 0;
        int found;
        while ((found = text.IndexOf(oldValue, index, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            sb.Append(text, index, found - index);
            sb.Append(newValue);
            index = found + oldValue.Length;
        }
        sb.Append(text, index, text.Length - index);
        return sb.ToString();
    }

    private static string ReplaceContentControl(string xml, string key, string escapedValue)
    {
        var escapedKey = Regex.Escape(key);
        var pattern = $"""
            <w:sdt\b(?:(?!</w:sdt>).)*?(?:<w:tag\s+w:val="{escapedKey}"\s*/>|<w:alias\s+w:val="{escapedKey}"\s/>)(?:(?!</w:sdt>).)*?</w:sdt>
            """;

        return Regex.Replace(xml, pattern, match =>
        {
            var valueWritten = false;
            return Regex.Replace(match.Value, "<w:t(\\s[^>]*)?>.*?</w:t>", textMatch =>
            {
                if (valueWritten)
                    return textMatch.Value.Replace(Regex.Match(textMatch.Value, ">.*<").Value, "><");

                valueWritten = true;
                var attrs = Regex.Match(textMatch.Value, "<w:t(\\s[^>]*)?>").Groups[1].Value;
                return $"<w:t{attrs}>{escapedValue}</w:t>";
            }, RegexOptions.Singleline);
        }, RegexOptions.Singleline);
    }

    private static bool IsEditableXml(string path) =>
        path.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
        (path.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) ||
        (path.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));

    private static string Escape(string text) =>
        new XElement("x", text).Value
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;");
}
