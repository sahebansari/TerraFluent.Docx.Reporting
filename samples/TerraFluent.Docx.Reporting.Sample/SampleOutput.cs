internal static class SampleOutput
{
    public static string EnsureDesktopSampleDocs()
    {
        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrWhiteSpace(desktopPath))
            desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");

        var sampleDocsDirectory = Path.Combine(desktopPath, "SampleDocs");
        Directory.CreateDirectory(sampleDocsDirectory);
        return sampleDocsDirectory;
    }
}
