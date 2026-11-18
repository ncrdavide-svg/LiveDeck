using System.Text.Json;

namespace LiveSlides.Core;

/// <summary>
/// Per-folder show settings, saved as "liveslides.json" inside the folder itself so the running
/// order travels with the files (copy the folder to another PC and the show is the same).
/// </summary>
public class ShowFile
{
    public const string FileName = "liveslides.json";

    /// <summary>Extension of a show saved on purpose with «Salva…» (same JSON, plus <see cref="Folder"/>).</summary>
    public const string ProjectExtension = ".liveslides";

    /// <summary>Only in .liveslides files: the show folder when it was saved (absolute path).</summary>
    public string? Folder { get; set; }

    /// <summary>File names (not paths) in running order.</summary>
    public List<string> Order { get; set; } = new();
    /// <summary>File names excluded from the show.</summary>
    public List<string> Disabled { get; set; } = new();
    /// <summary>Tappo image: a bare file name when it lives in the show folder, else a full path.</summary>
    public string? Tappo { get; set; }
    public bool TappoAtStart { get; set; } = true;
    public bool TappoBetween { get; set; } = true;
    public bool TappoAtEnd { get; set; } = true;
    public int FadeMs { get; set; } = 300;

    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public static ShowFile Load(string folder)
    {
        try
        {
            var path = Path.Combine(folder, FileName);
            if (File.Exists(path))
                return JsonSerializer.Deserialize<ShowFile>(File.ReadAllText(path)) ?? new ShowFile();
        }
        catch
        {
            // A corrupt settings file must not stop the folder from loading: start from defaults.
        }
        return new ShowFile();
    }

    public void Save(string folder)
    {
        // The folder's own copy never carries a folder path: it is the folder.
        var copy = (ShowFile)MemberwiseClone();
        copy.Folder = null;
        File.WriteAllText(Path.Combine(folder, FileName), JsonSerializer.Serialize(copy, Options));
    }

    /// <summary>Writes a .liveslides project file anywhere the user chooses.</summary>
    public void SaveProject(string path) => File.WriteAllText(path, JsonSerializer.Serialize(this, Options));

    /// <summary>Reads a .liveslides project file. Unlike <see cref="Load"/>, a corrupt file throws:
    /// the user asked for this exact file and must be told it can't be read.</summary>
    public static ShowFile LoadProject(string path)
        => JsonSerializer.Deserialize<ShowFile>(File.ReadAllText(path))
           ?? throw new InvalidDataException("Il file è vuoto.");

    /// <summary>
    /// Where the presentations of a project are. A project saved inside its show folder keeps
    /// working when the whole folder is copied to another PC (or another drive letter), so the
    /// project file's own folder wins when it holds the show's files; otherwise the saved path.
    /// Null when neither is there.
    /// </summary>
    public static string? ResolveProjectFolder(string projectPath, ShowFile show)
    {
        var own = Path.GetDirectoryName(Path.GetFullPath(projectPath));
        if (own is not null && show.Order.Any(name => File.Exists(Path.Combine(own, name)))) return own;
        if (!string.IsNullOrWhiteSpace(show.Folder) && Directory.Exists(show.Folder)) return show.Folder;
        if (own is not null && show.Order.Count == 0 && string.IsNullOrWhiteSpace(show.Folder)) return own;
        return null;
    }

    /// <summary>
    /// Running order for the files currently in the folder: the saved order first (dropping
    /// files that no longer exist), then any new file in alphabetical order at the end.
    /// </summary>
    public static List<string> ApplyOrder(IEnumerable<string> namesInFolder, IReadOnlyList<string> savedOrder)
    {
        var present = new HashSet<string>(namesInFolder, StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        foreach (var name in savedOrder)
            if (present.Remove(name)) result.Add(name);
        result.AddRange(present.OrderBy(n => n, StringComparer.OrdinalIgnoreCase));
        return result;
    }

    public static string? ResolveTappoPath(string folder, string? tappo)
    {
        if (string.IsNullOrWhiteSpace(tappo)) return null;
        return Path.IsPathRooted(tappo) ? tappo : Path.Combine(folder, tappo);
    }

    public static string? MakeTappoReference(string folder, string? fullPath)
    {
        if (string.IsNullOrWhiteSpace(fullPath)) return null;
        var dir = Path.GetDirectoryName(fullPath);
        return string.Equals(dir?.TrimEnd('\\', '/'), folder.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase)
            ? Path.GetFileName(fullPath)
            : fullPath;
    }
}
