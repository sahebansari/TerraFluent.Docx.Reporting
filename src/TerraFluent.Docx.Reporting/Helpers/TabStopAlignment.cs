namespace TerraFluent.Docx.Reporting;

/// <summary>
/// How text is aligned relative to a custom tab stop, set via <see cref="Infra.ITextDescriptor.TabStop"/>.
/// </summary>
public enum TabStopAlignment
{
    /// <summary>Text starts at the tab stop and flows right.</summary>
    Left,
    /// <summary>Text is centered on the tab stop.</summary>
    Center,
    /// <summary>Text ends at the tab stop, flowing left from it.</summary>
    Right,
    /// <summary>Numbers align on their decimal point at the tab stop.</summary>
    Decimal
}
