using LiveSlides.Core;
using Xunit;

public class ShowFileTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "LiveSlidesTests-" + Guid.NewGuid().ToString("N"));

    public ShowFileTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    private string MakeFolder(string name, params string[] files)
    {
        var folder = Path.Combine(_root, name);
        Directory.CreateDirectory(folder);
        foreach (var f in files) File.WriteAllText(Path.Combine(folder, f), "x");
        return folder;
    }

    [Fact]
    public void ProjectRoundTripsEverySetting()
    {
        var path = Path.Combine(_root, "show" + ShowFile.ProjectExtension);
        new ShowFile
        {
            Folder = @"D:\Show", Order = { "b.pptx", "a.pdf" }, Disabled = { "a.pdf" },
            Tappo = "logo.png", TappoAtStart = false, TappoBetween = true, TappoAtEnd = false, FadeMs = 750
        }.SaveProject(path);

        var back = ShowFile.LoadProject(path);
        Assert.Equal(@"D:\Show", back.Folder);
        Assert.Equal(new[] { "b.pptx", "a.pdf" }, back.Order);
        Assert.Equal(new[] { "a.pdf" }, back.Disabled);
        Assert.Equal("logo.png", back.Tappo);
        Assert.False(back.TappoAtStart);
        Assert.True(back.TappoBetween);
        Assert.False(back.TappoAtEnd);
        Assert.Equal(750, back.FadeMs);
    }

    [Fact]
    public void FolderCopyDoesNotStoreTheFolderPath()
    {
        var folder = MakeFolder("show");
        new ShowFile { Folder = folder }.Save(folder);
        Assert.Null(ShowFile.Load(folder).Folder);
    }

    [Fact]
    public void CorruptProjectThrows()
    {
        var path = Path.Combine(_root, "bad" + ShowFile.ProjectExtension);
        File.WriteAllText(path, "{ not json");
        Assert.ThrowsAny<Exception>(() => ShowFile.LoadProject(path));
    }

    [Fact]
    public void ProjectInsideACopiedFolderUsesThatFolder()
    {
        // Saved on D:, the folder was then copied elsewhere together with the project file.
        var copied = MakeFolder("copied", "a.pdf", "b.pptx");
        var project = Path.Combine(copied, "show" + ShowFile.ProjectExtension);
        var show = new ShowFile { Folder = @"Z:\does-not-exist", Order = { "a.pdf", "b.pptx" } };
        Assert.Equal(copied, ShowFile.ResolveProjectFolder(project, show));
    }

    [Fact]
    public void ProjectSavedElsewhereUsesTheSavedFolder()
    {
        var folder = MakeFolder("decks", "a.pdf");
        var project = Path.Combine(MakeFolder("projects"), "show" + ShowFile.ProjectExtension);
        Assert.Equal(folder, ShowFile.ResolveProjectFolder(project, new ShowFile { Folder = folder, Order = { "a.pdf" } }));
    }

    [Fact]
    public void MissingFolderGivesNull()
    {
        var project = Path.Combine(MakeFolder("projects"), "show" + ShowFile.ProjectExtension);
        Assert.Null(ShowFile.ResolveProjectFolder(project, new ShowFile { Folder = @"Z:\nope", Order = { "a.pdf" } }));
    }
}
