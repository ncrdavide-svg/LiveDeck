using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace LiveDeck.App;

public partial class App : Application
{
    // LIVEDECK_DATA overrides the folder: lets a second, fully separate instance (own settings,
    // own remote port) run for testing next to the one being used for a show.
    public static string DataFolder { get; } =
        Environment.GetEnvironmentVariable("LIVEDECK_DATA") is { Length: > 0 } custom
            ? custom
            : DefaultDataFolder();

    /// <summary>
    /// %LOCALAPPDATA%\LiveDeck. The app was called LiveSlides until 1.4: the first start after the
    /// update moves the old data folder here, so settings, monitor choices and the LibreOffice
    /// profiles (which avoid the one-time printer message) are kept.
    /// </summary>
    private static string DefaultDataFolder()
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var folder = Path.Combine(local, "LiveDeck");
        var legacy = Path.Combine(local, "LiveSlides");
        if (!Directory.Exists(legacy)) return folder;
        try
        {
            if (!Directory.Exists(folder))
            {
                Directory.Move(legacy, folder);
                return folder;
            }
            // A LiveDeck folder can already be there (only a cache, from a test build): move over
            // each old file and folder it does not have yet, so settings.json is never lost.
            foreach (var entry in Directory.EnumerateFileSystemEntries(legacy))
            {
                var target = Path.Combine(folder, Path.GetFileName(entry));
                if (File.Exists(target) || Directory.Exists(target)) continue;
                if (Directory.Exists(entry)) Directory.Move(entry, target);
                else File.Move(entry, target);
            }
            if (!Directory.EnumerateFileSystemEntries(legacy).Any()) Directory.Delete(legacy);
        }
        catch
        {
            // Something still holds the old folder (an old LiveSlides open?): keep going with what moved.
        }
        return folder;
    }

    /// <summary>%LOCALAPPDATA%\LiveDeck\errors.log — every unexpected error is appended here.</summary>
    public static string ErrorLogPath { get; } = Path.Combine(DataFolder, "errors.log");

    protected override void OnStartup(StartupEventArgs e)
    {
        // A show app must never close in the middle of a show because of one bad file:
        // unexpected errors are logged and reported in the status bar, and the app keeps running.
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            LogError(args.Exception);
            args.SetObserved();
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex) LogError(ex);
        };
        base.OnStartup(e);
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        LogError(e.Exception);
        e.Handled = true;

        var message = $"Errore: {e.Exception.Message} (dettagli in {ErrorLogPath})";
        if (MainWindow?.DataContext is MainViewModel viewModel)
            viewModel.StatusMessage = message;
        else
            MessageBox.Show(message, "LiveDeck", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    public static void LogError(Exception exception)
    {
        try
        {
            Directory.CreateDirectory(DataFolder);
            File.AppendAllText(ErrorLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {exception}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
            // logging must never throw
        }
    }
}
