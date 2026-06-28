namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Builds a vertical stack of content inside an <see cref="IContainer.Column"/> call.
/// </summary>
public interface IColumnDescriptor
{
    /// <summary>Sets the vertical spacing, in points, inserted between each item added to the column.</summary>
    IColumnDescriptor Spacing(float points);

    /// <summary>Adds a new content item to the column and returns its container.</summary>
    IContainer Item();
}
