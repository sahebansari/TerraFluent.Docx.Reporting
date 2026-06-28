namespace TerraFluent.Docx.Reporting.Infra;

/// <summary>
/// A reusable piece of content that can be composed into any <see cref="IContainer"/> via
/// <see cref="IContainer.Component"/>. Implement this to share repeated layout (e.g. a letterhead
/// or a callout box) across documents or pages.
/// </summary>
public interface IComponent
{
    /// <summary>Adds this component's content to the given container.</summary>
    /// <param name="container">The container to add content to.</param>
    void Compose(IContainer container);
}
