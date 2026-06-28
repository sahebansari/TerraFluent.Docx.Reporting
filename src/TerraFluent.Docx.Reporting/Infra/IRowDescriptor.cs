namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Builds a horizontal layout row inside an <see cref="IContainer.Row"/> call. A row is made up of
/// side-by-side item containers sized as relative, automatic, or fixed-width columns.
/// </summary>
public interface IRowDescriptor
{
    /// <summary>Sets the horizontal spacing, in points, inserted between each item added to the row.</summary>
    IRowDescriptor Spacing(float points);

    /// <summary>Adds an item container whose width is proportional to other relative items in the row.</summary>
    /// <param name="size">The item's relative size weight compared to other <see cref="RelativeItem"/> calls in the same row.</param>
    IContainer RelativeItem(float size = 1);

    /// <summary>Adds an item container sized to fit its content.</summary>
    IContainer AutoItem();

    /// <summary>Adds an item container with a fixed width.</summary>
    /// <param name="widthPoints">The item's width in points.</param>
    IContainer ConstantItem(float widthPoints);
}
