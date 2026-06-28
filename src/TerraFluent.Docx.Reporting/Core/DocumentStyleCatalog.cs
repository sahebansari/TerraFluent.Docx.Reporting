using System.Text;
using TerraFluent.Docx.Reporting.Core.Elements;

namespace TerraFluent.Docx.Reporting.Core;

internal sealed class DocumentStyleCatalog
{
    private readonly Dictionary<string, ParagraphStyleDefinition> _paragraphStyles = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, TableStyleDefinition> _tableStyles = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyCollection<ParagraphStyleDefinition> ParagraphStyles => _paragraphStyles.Values;
    public IReadOnlyCollection<TableStyleDefinition> TableStyles => _tableStyles.Values;

    public void AddParagraphStyle(string name, TextElement style)
    {
        _paragraphStyles[name] = new ParagraphStyleDefinition(name, StyleId(name), style);
    }

    public void AddTableStyle(string name, TableElement style)
    {
        _tableStyles[name] = new TableStyleDefinition(name, StyleId(name), style);
    }

    public bool TryGetTableStyle(string name, out TableStyleDefinition style) => _tableStyles.TryGetValue(name, out style!);

    public static string StyleId(string name)
    {
        var sb = new StringBuilder();
        foreach (var ch in name)
        {
            if (char.IsLetterOrDigit(ch))
                sb.Append(ch);
        }

        return sb.Length == 0 ? "CustomStyle" : sb.ToString();
    }
}

internal sealed record ParagraphStyleDefinition(string Name, string StyleId, TextElement Style);

internal sealed record TableStyleDefinition(string Name, string StyleId, TableElement Style);
