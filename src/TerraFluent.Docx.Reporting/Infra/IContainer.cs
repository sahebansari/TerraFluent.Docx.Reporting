namespace TerraFluent.Docx.Reporting.Infra;

public interface IContainer
{
    IContainer H1(string text);
    IContainer H2(string text);
    IContainer H3(string text);
    IContainer H4(string text);
    IContainer H5(string text);
    IContainer H6(string text);
    IContainer Column(Action<IColumnDescriptor> configure);
    IContainer Row(Action<IRowDescriptor> configure);
    IContainer Text(string text, Action<ITextDescriptor>? configure = null);
    IContainer Text(Action<ITextDescriptor> configure);
    IContainer Hyperlink(string text, string url, Action<ITextDescriptor>? configure = null);
    IContainer Bookmark(string name);
    IContainer Bookmark(string name, string text, Action<ITextDescriptor>? configure = null);
    IContainer TableOfContents(string title = "Contents", int minLevel = 1, int maxLevel = 3);
    IContainer BulletList(Action<IListDescriptor> configure);
    IContainer NumberedList(Action<IListDescriptor> configure);
    IContainer Table(Action<ITableDescriptor> configure);
    IContainer Chart(Action<IChartDescriptor> configure);
    IContainer Image(string filePath, float? width = null);
    IContainer Image(string filePath, Action<IImageDescriptor> configure);
    IContainer Image(byte[] imageBytes, string fileName, Action<IImageDescriptor>? configure = null);
    IContainer Image(Stream imageStream, string fileName, Action<IImageDescriptor>? configure = null);
    IContainer Line();
    IContainer PageBreak();
    IContainer Component(IComponent component);
}
