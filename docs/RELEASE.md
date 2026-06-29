# Release And Compatibility Checklist

[Documentation Home](README.md) | [Getting Started](GETTING_STARTED.md) | [API Reference](API.md) | [Troubleshooting](TROUBLESHOOTING.md)

TerraFluent.Docx.Reporting uses semantic versioning:

- `MAJOR`: breaking public API or generated document contract changes.
- `MINOR`: backward-compatible features.
- `PATCH`: bug fixes and compatibility repairs.
- Preview releases should use a prerelease suffix, for example `0.2.0-preview.1`.

The package version is stored in `src/TerraFluent.Docx.Reporting/TerraFluent.Docx.Reporting.csproj` as `VersionPrefix`.

Before publishing, make sure:

- `CHANGELOG.md` has an entry for the version.
- `LICENSE.txt` contains the MIT license and current copyright holder.
- The package metadata in the project file has a real project/repository URL.
- The solution builds with warnings treated as errors.
- A nuget.org Trusted Publishing policy exists for the GitHub Actions release environment.
- The GitHub repository has a `NUGET_USER` secret containing the nuget.org username for Trusted Publishing.

## Local Release Build

```powershell
dotnet clean TerraFluent.Docx.Reporting.sln -c Release
dotnet test TerraFluent.Docx.Reporting.sln -c Release
dotnet pack src\TerraFluent.Docx.Reporting\TerraFluent.Docx.Reporting.csproj -c Release -o artifacts\nuget
```

Before publishing, inspect the package:

```powershell
tar -tf artifacts\nuget\TerraFluent.Docx.Reporting.*.nupkg
```

The package should include:

- `lib/netstandard2.0/TerraFluent.Docx.Reporting.dll`
- `lib/netstandard2.0/TerraFluent.Docx.Reporting.xml`
- `lib/net10.0/TerraFluent.Docx.Reporting.dll`
- `lib/net10.0/TerraFluent.Docx.Reporting.xml`
- `README.md`
- `CHANGELOG.md`
- `LICENSE.txt`
- `docs/*.md` (every file in the `docs` folder, including `README.md`, `API.md`, `GETTING_STARTED.md`, `CORE_CONCEPTS.md`, `FEATURES.md`, `SAMPLES.md`, `TROUBLESHOOTING.md`, and `RELEASE.md`)

Confirm the package metadata:

- Package ID: `TerraFluent.Docx.Reporting`
- License expression: `MIT`
- Symbols package: `.snupkg`
- Repository URL and project URL point to the public source repository.

## Compatibility Smoke Checks

The automated test suite validates generated DOCX packages with the Open XML SDK. CI also builds on Windows and Linux, runs the sample project, performs a packed-package consumption smoke test, and converts sample documents with LibreOffice on Linux.

Application-level compatibility should also be checked before a release:

```powershell
dotnet run --project samples\TerraFluent.Docx.Reporting.Sample\TerraFluent.Docx.Reporting.Sample.csproj
```

Then manually verify the generated files in the applications you support:

- Open and save the sample documents in Microsoft Word.
- Open or convert the sample documents in LibreOffice.
- Confirm there are no repair prompts, missing images, broken charts, or layout regressions.

Use CI machines with Word or LibreOffice installed when these checks need to be automated.

## Release Notes

Each release should include:

- New public APIs.
- Generated DOCX compatibility fixes.
- Known limitations.
- Migration notes for any breaking changes.

## Publish

Publishing can be done from GitHub Actions by pushing a version tag such as `v1.2.1`, or by running the `CI` workflow manually with `publish=true`.

GitHub Actions publishing uses nuget.org Trusted Publishing instead of a long-lived API key. Configure the nuget.org Trusted Publishing policy for:

- Repository owner: `sahebansari`
- Repository: `TerraFluent.Docx.Reporting`
- Workflow: `ci.yml`
- Environment: `release`

To publish locally, pack first and then push:

```powershell
dotnet nuget push artifacts\nuget\TerraFluent.Docx.Reporting.<version>.nupkg --source https://api.nuget.org/v3/index.json --api-key <key>
```

Publish symbols if configured by the NuGet source:

```powershell
dotnet nuget push artifacts\nuget\TerraFluent.Docx.Reporting.<version>.snupkg --source https://api.nuget.org/v3/index.json --api-key <key>
```
