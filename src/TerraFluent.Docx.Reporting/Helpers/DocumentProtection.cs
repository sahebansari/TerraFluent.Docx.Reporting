namespace TerraFluent.Docx.Reporting;

/// <summary>
/// The editing restriction applied via <see cref="Infra.IDocumentContainer.RestrictEditing"/> -
/// Word's "Restrict Editing" feature. This limits what a reader can change through Word's UI;
/// it is not encryption, and the underlying file content remains readable.
/// </summary>
public enum DocumentProtection
{
    /// <summary>The document is read-only; no edits are allowed.</summary>
    ReadOnly,
    /// <summary>Readers may only insert comments.</summary>
    CommentsOnly,
    /// <summary>Edits are allowed but are always recorded as tracked changes.</summary>
    TrackedChangesOnly,
    /// <summary>Readers may only fill in form fields.</summary>
    FormsOnly
}
