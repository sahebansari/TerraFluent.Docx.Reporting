# Contributing To TerraFluent.Docx.Reporting

Thanks for helping improve TerraFluent.Docx.Reporting. This project is a fluent .NET library for generating `.docx` files directly with Open XML, so contributions should preserve predictable output, strong compatibility with Word processors, and a clean public API.

## Ways To Contribute

- Report bugs with a minimal reproduction and the generated `.docx` behavior you expected.
- Improve documentation, samples, and troubleshooting notes.
- Add focused tests for DOCX generation, validation, and packaging behavior.
- Propose API improvements before doing large implementation work.

## Development Setup

Install the .NET 10 SDK, then restore and build the solution:

```powershell
dotnet restore TerraFluent.Docx.Reporting.sln
dotnet build TerraFluent.Docx.Reporting.sln
```

Run the full test suite before opening a pull request:

```powershell
dotnet test TerraFluent.Docx.Reporting.sln
```

Run the sample project when your change affects generated output:

```powershell
dotnet run --project samples\TerraFluent.Docx.Reporting.Sample\TerraFluent.Docx.Reporting.Sample.csproj
```

Generated sample documents are written to `Desktop\SampleDocs`.

## Code Guidelines

- Keep public APIs fluent, discoverable, and consistent with the existing descriptors.
- Prefer explicit, validation-friendly Open XML output over shortcuts that only work in one application.
- Keep changes small and focused. Avoid unrelated formatting or refactoring in feature and bug-fix pull requests.
- Keep nullable reference types clean and follow the existing C# style in nearby files.
- Add XML documentation for new public APIs where it helps package consumers.

## Tests And Compatibility

The automated tests validate generated packages with the Open XML SDK. Add or update tests when changing:

- Public document-building APIs.
- Open XML serialization.
- Image, chart, table, header, footer, or template behavior.
- NuGet packaging metadata or release output.

For layout-sensitive changes, also open the generated sample files in Microsoft Word and LibreOffice when possible. Note any manual compatibility checks in the pull request.

## Pull Request Checklist

Before submitting a pull request:

- Run `dotnet test TerraFluent.Docx.Reporting.sln`.
- Update samples or docs for user-facing behavior changes.
- Update `CHANGELOG.md` for notable changes.
- Confirm generated documents do not trigger repair prompts in supported applications.
- Keep the PR description clear about the problem, the fix, and any compatibility risk.

## Issues

When opening an issue, include:

- TerraFluent.Docx.Reporting version or commit.
- Operating system and .NET SDK version.
- A minimal code sample.
- The expected result and actual result.
- Any relevant document repair message, validation error, or viewer-specific behavior.

For vulnerabilities, please follow [SECURITY.md](SECURITY.md) instead of opening a public issue.
