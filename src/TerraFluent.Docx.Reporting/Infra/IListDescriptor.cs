namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// Builds a bulleted or numbered list inside an <see cref="IContainer.BulletList"/> or
/// <see cref="IContainer.NumberedList"/> call.
/// </summary>
public interface IListDescriptor
{
    /// <summary>Sets a custom marker (e.g. a bullet character) for a given list level.</summary>
    /// <param name="marker">The marker text or character to display before items at this level.</param>
    /// <param name="level">The list level, where 0 is the top level.</param>
    /// <param name="fontFamily">An optional font family for the marker, useful for symbol fonts.</param>
    IListDescriptor Marker(string marker, int level = 0, string? fontFamily = null);

    /// <summary>Adds a plain-text item at the top list level (level 0).</summary>
    /// <param name="text">The item text.</param>
    /// <param name="configure">Optional formatting for the item's paragraph.</param>
    IListDescriptor Item(string text, Action<ITextDescriptor>? configure = null);

    /// <summary>Adds a plain-text item at a specific list level, for nested lists.</summary>
    /// <param name="text">The item text.</param>
    /// <param name="level">The list level, where 0 is the top level.</param>
    /// <param name="configure">Optional formatting for the item's paragraph.</param>
    IListDescriptor Item(string text, int level, Action<ITextDescriptor>? configure = null);

    /// <summary>Adds a rich-text item at the top list level (level 0), built from multiple formatted runs.</summary>
    /// <param name="configure">Builds the item's runs and formatting. Cannot be null.</param>
    IListDescriptor Item(Action<ITextDescriptor> configure);

    /// <summary>Adds a rich-text item at a specific list level, built from multiple formatted runs.</summary>
    /// <param name="level">The list level, where 0 is the top level.</param>
    /// <param name="configure">Builds the item's runs and formatting. Cannot be null.</param>
    IListDescriptor Item(int level, Action<ITextDescriptor> configure);
}
