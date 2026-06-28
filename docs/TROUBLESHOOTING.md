# Troubleshooting

[Documentation Home](README.md) | [Getting Started](GETTING_STARTED.md) | [Feature Guide](FEATURES.md) | [API Reference](API.md)

## Word Shows A Repair Prompt

Run the automated tests and inspect the generated document with Word.

```powershell
dotnet test TerraFluent.Docx.Reporting.sln
```

Common causes:

- Invalid color values. Use six-character hex strings such as `1F4E79`, not `#1F4E79`.
- Invalid page number formats. `PageNumberFormat` writes the OOXML value as-is. Use values like `decimal`, `lowerRoman`, `upperRoman`, `lowerLetter`, or `upperLetter`.
- Mixed chart types in one chart. All chart series must use the same kind.
- Multiple series in a pie or doughnut chart. These chart types support one series.

## Images Do Not Appear

Check the path or byte data supplied to the image API.

```csharp
page.Content().Image("logo.png", image => image
    .Width(120)
    .AltText("Company logo"));
```

Tips:

- Use an absolute path when the current working directory is unclear.
- Ensure file names have supported extensions such as `.png`, `.jpg`, or `.jpeg`.
- Use `AltText` for accessibility and easier inspection.
- Use `MaxWidth` when documents may receive images of unknown size.

## Page Background Does Not Change Per Section

Word stores page background as a document-wide setting. If multiple sections call `Background`, the first non-empty value wins.

Use section-specific watermarks or shaded containers when you need visual differences per section.

## Table Of Contents Does Not Show Page Numbers Immediately

The TOC is emitted as a Word field. Word normally updates fields when you open the document or explicitly refresh fields.

In Word, select the table of contents and choose update field.

## Layout Differs Between Word And LibreOffice

Open XML rendering can differ between applications. Before publishing templates or report layouts:

- Test in Microsoft Word.
- Test in LibreOffice if your users rely on it.
- Avoid overly tight fixed-width tables.
- Prefer relative columns for responsive report tables.
- Keep images within page margins with `MaxWidth`.

## NuGet Package Missing Docs Or License

Pack the project and inspect the package:

```powershell
dotnet pack src\TerraFluent.Docx.Reporting\TerraFluent.Docx.Reporting.csproj -c Release -o artifacts\nuget
tar -tf artifacts\nuget\TerraFluent.Docx.Reporting.*.nupkg
```

The package should include:

- `README.md`
- `CHANGELOG.md`
- `LICENSE.txt`
- `lib/netstandard2.0/TerraFluent.Docx.Reporting.dll`
- `lib/netstandard2.0/TerraFluent.Docx.Reporting.xml`
- `lib/net10.0/TerraFluent.Docx.Reporting.dll`
- `lib/net10.0/TerraFluent.Docx.Reporting.xml`

See [Release And Publishing](RELEASE.md) for the full checklist.

