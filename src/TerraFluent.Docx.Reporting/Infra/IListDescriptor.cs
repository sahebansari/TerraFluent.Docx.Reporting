namespace TerraFluent.Docx.Reporting.Infra;

public interface IListDescriptor
{
    IListDescriptor Marker(string marker, int level = 0, string? fontFamily = null);
    IListDescriptor Item(string text, Action<ITextDescriptor>? configure = null);
    IListDescriptor Item(string text, int level, Action<ITextDescriptor>? configure = null);
    IListDescriptor Item(Action<ITextDescriptor> configure);
    IListDescriptor Item(int level, Action<ITextDescriptor> configure);
}
