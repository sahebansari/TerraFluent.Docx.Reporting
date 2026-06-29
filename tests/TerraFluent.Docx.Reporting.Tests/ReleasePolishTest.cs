using System.Reflection;
using System.Xml.Linq;

namespace TerraFluent.Docx.Reporting.Tests;

public class ReleasePolishTest
{
    [Fact]
    public void NuGetPackageMetadata_IsConfigured()
    {
        var projectPath = Path.Combine(RepositoryRoot(), "src", "TerraFluent.Docx.Reporting", "TerraFluent.Docx.Reporting.csproj");
        var project = XDocument.Load(projectPath);
        var properties = project.Root!.Elements("PropertyGroup").Elements()
            .GroupBy(e => e.Name.LocalName)
            .ToDictionary(g => g.Key, g => g.Last().Value);

        Assert.Equal("TerraFluent.Docx.Reporting", properties["PackageId"]);
        Assert.Equal("TerraFluent.Docx.Reporting", properties["Title"]);
        Assert.Matches(@"^\d+\.\d+\.\d+", properties["VersionPrefix"]);
        Assert.False(string.IsNullOrWhiteSpace(properties["Description"]));
        Assert.Equal("MIT", properties["PackageLicenseExpression"]);
        Assert.Equal("README.md", properties["PackageReadmeFile"]);
        Assert.Equal("false", properties["PackageRequireLicenseAcceptance"]);
        Assert.Equal("git", properties["RepositoryType"]);
        Assert.StartsWith("https://", properties["RepositoryUrl"]);
        Assert.StartsWith("https://", properties["PackageProjectUrl"]);
        Assert.Equal("true", properties["GenerateDocumentationFile"]);
        Assert.Equal("snupkg", properties["SymbolPackageFormat"]);
        Assert.Equal("true", properties["PublishRepositoryUrl"]);

        var packedFiles = project.Root.Elements("ItemGroup").Elements("None")
            .Where(e => string.Equals(e.Attribute("Pack")?.Value, "true", StringComparison.OrdinalIgnoreCase))
            .Select(e => GetMsBuildFileName(e.Attribute("Include")?.Value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        Assert.Contains("README.md", packedFiles);
        Assert.Contains("CHANGELOG.md", packedFiles);
        Assert.Contains("LICENSE.txt", packedFiles);

        Assert.Contains(project.Root.Elements("ItemGroup").Elements("None"), e =>
            string.Equals(e.Attribute("Include")?.Value, @"..\..\docs\**\*.md", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(e.Attribute("Pack")?.Value, "true", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(e.Attribute("PackagePath")?.Value, @"docs\", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ReleaseDocumentation_Exists()
    {
        var root = RepositoryRoot();

        Assert.Contains("dotnet add package TerraFluent.Docx.Reporting", File.ReadAllText(Path.Combine(root, "README.md")));
        Assert.Contains("MIT License", File.ReadAllText(Path.Combine(root, "LICENSE.txt")));
        Assert.DoesNotContain("[year]", File.ReadAllText(Path.Combine(root, "LICENSE.txt")));
        Assert.Contains("## 1.2.0", File.ReadAllText(Path.Combine(root, "CHANGELOG.md")));
        Assert.Contains("Documentation home and main menu", File.ReadAllText(Path.Combine(root, "README.md")));
        Assert.Contains("Main Menu", File.ReadAllText(Path.Combine(root, "docs", "README.md")));
        Assert.Contains("Create Your First Document", File.ReadAllText(Path.Combine(root, "docs", "GETTING_STARTED.md")));
        Assert.Contains("Reusable Components", File.ReadAllText(Path.Combine(root, "docs", "CORE_CONCEPTS.md")));
        Assert.Contains("Feature Guide", File.ReadAllText(Path.Combine(root, "docs", "FEATURES.md")));
        Assert.Contains("API Reference", File.ReadAllText(Path.Combine(root, "docs", "API.md")));
        Assert.Contains("Sample Files", File.ReadAllText(Path.Combine(root, "docs", "SAMPLES.md")));
        Assert.Contains("Word Shows A Repair Prompt", File.ReadAllText(Path.Combine(root, "docs", "TROUBLESHOOTING.md")));
        Assert.Contains("semantic versioning", File.ReadAllText(Path.Combine(root, "docs", "RELEASE.md")));
        Assert.Contains("Trusted Publishing", File.ReadAllText(Path.Combine(root, "docs", "RELEASE.md")));
        Assert.Contains("NUGET_USER", File.ReadAllText(Path.Combine(root, "docs", "RELEASE.md")));
        Assert.Contains("Open and save the sample documents in Microsoft Word", File.ReadAllText(Path.Combine(root, "docs", "RELEASE.md")));
    }

    [Fact]
    public void GitHubWorkflow_CanBuildPackAndPublish()
    {
        var workflowPath = Path.Combine(RepositoryRoot(), ".github", "workflows", "ci.yml");
        var workflow = File.ReadAllText(workflowPath);

        Assert.Contains("dotnet test TerraFluent.Docx.Reporting.sln", workflow);
        Assert.Contains("dotnet pack src\\TerraFluent.Docx.Reporting\\TerraFluent.Docx.Reporting.csproj", workflow);
        Assert.Contains("NuGet/login@v1", workflow);
        Assert.Contains("refs/heads/main", workflow);
        Assert.Contains("refs/heads/master", workflow);
        Assert.Contains("refs/tags/v", workflow);
        Assert.Contains("id-token: write", workflow);
        Assert.Contains("environment: release", workflow);
        Assert.Contains("secrets.NUGET_USER", workflow);
        Assert.DoesNotContain("secrets.NUGET_API_KEY", workflow);
    }

    [Fact]
    public void XmlDocumentationFile_IsGeneratedForPublicApi()
    {
        var assemblyPath = typeof(Document).Assembly.Location;
        var xmlPath = Path.ChangeExtension(assemblyPath, ".xml");

        Assert.True(File.Exists(xmlPath), $"Expected XML documentation file next to {Path.GetFileName(assemblyPath)}.");
        Assert.Contains("Represents a TerraFluent.Docx.Reporting document", File.ReadAllText(xmlPath));
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "TerraFluent.Docx.Reporting.sln")))
            directory = directory.Parent;

        return directory?.FullName ?? throw new DirectoryNotFoundException("Could not find TerraFluent.Docx.Reporting.sln.");
    }

    private static string? GetMsBuildFileName(string? path)
    {
        return path?.Split('\\', '/').LastOrDefault();
    }
}
