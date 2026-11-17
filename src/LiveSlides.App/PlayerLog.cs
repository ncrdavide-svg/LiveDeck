using System.IO;

namespace LiveSlides.App;

/// <summary>
/// One line per operation of the presentation programs in player.log (data folder): when a
/// presentation does not go on air as expected, this says what PowerPoint or LibreOffice did.
/// Restarted at each run of the app, so it stays small.
/// </summary>
internal static class PlayerLog
{
    private static readonly object Gate = new();
    private static bool _started;

    public static string FilePath => Path.Combine(App.DataFolder, "player.log");

    public static void Write(string message)
    {
        try
        {
            lock (Gate)
            {
                if (!_started) { _started = true; File.WriteAllText(FilePath, ""); }
                File.AppendAllText(FilePath, $"{DateTime.Now:HH:mm:ss.fff} {message}{Environment.NewLine}");
            }
        }
        catch { /* diagnostics only */ }
    }
}
