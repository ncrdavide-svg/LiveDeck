using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace LiveSlides.App;

/// <summary>
/// Plays presentations with Microsoft PowerPoint itself, through COM automation: the file runs
/// exactly as its author designed it. All COM calls happen on one dedicated STA thread.
///
/// Written from the documented object model and not yet confirmed on a PC with PowerPoint: every
/// step is written to player.log (data folder) so a failed try says what happened.
/// </summary>
public sealed class PowerPointPlayer : IPresentationPlayer
{
    // MsoTriState / PowerPoint enums (late binding: no interop assembly to ship).
    private const int MsoTrue = -1, MsoFalse = 0;
    private const int PpShowTypeSpeaker = 1;
    private const int PpSlideShowPaused = 2, PpSlideShowRunning = 1, PpSlideShowDone = 5;

    private readonly Type _type;
    private readonly BlockingCollection<Action> _work = new();
    private readonly Thread _thread;
    private dynamic? _app;
    private dynamic? _presentation;
    private dynamic? _showWindow;
    // Files stay loaded (without a window) for the whole show: opening is the slow part.
    private readonly Dictionary<string, object> _loaded = new(StringComparer.OrdinalIgnoreCase);

    public PowerPointPlayer(Type powerPointType)
    {
        _type = powerPointType;
        _thread = new Thread(() => { foreach (var action in _work.GetConsumingEnumerable()) action(); })
        { IsBackground = true, Name = "PowerPoint show" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public string EngineName => "PowerPoint";
    public IntPtr Window { get; private set; }
    public IntPtr PresenterWindow { get; private set; }

    private static void Log(string message) => PlayerLog.Write("PowerPoint " + message);

    private Task<T> OnSta<T>(Func<T> work)
    {
        var done = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _work.Add(() =>
        {
            try { done.SetResult(Retry(work)); }
            catch (Exception ex) { done.SetException(ex); }
        });
        return done.Task;
    }

    /// <summary>PowerPoint rejects calls while it is busy (opening a file, mid-transition): try again shortly.</summary>
    private static T Retry<T>(Func<T> work)
    {
        for (var attempt = 0; ; attempt++)
        {
            try { return work(); }
            catch (COMException ex) when (attempt < 20 && (uint)ex.HResult is 0x80010001 or 0x8001010A)
            {
                Thread.Sleep(100);
            }
        }
    }

    /// <summary>
    /// The PowerPoint instance, started if needed and checked alive. PowerPoint runs as a single
    /// instance per PC, and the PDF conversion closes it when it has nothing open: an instance
    /// kept from before may be gone, and one asked for just as it closes answers with errors for a
    /// moment. Both cases get a fresh, working instance here.
    /// </summary>
    private dynamic Instance(ScreenBounds bounds)
    {
        if (_app is not null)
        {
            try { _ = (string)_app.Version; return _app; }
            catch (Exception ex)
            {
                Log($"instance gone ({ex.Message}): starting a new one");
                Forget();
            }
        }

        PreferMonitor(bounds);
        Exception? last = null;
        for (var attempt = 1; attempt <= 10; attempt++)
        {
            try
            {
                dynamic app = Activator.CreateInstance(_type)!;
                var version = (string)app.Version; // fails if it is the one that is closing
                Log($"instance ready (PowerPoint {version}, attempt {attempt})");
                _app = app;
                return app;
            }
            catch (Exception ex)
            {
                last = ex;
                Log($"instance not ready (attempt {attempt}): {ex.Message}");
                Thread.Sleep(500);
            }
        }
        throw new InvalidOperationException("PowerPoint non risponde: " + last?.Message, last);
    }

    /// <summary>Drops every reference to a PowerPoint instance that is no longer there.</summary>
    private void Forget()
    {
        _app = null;
        _presentation = null;
        _showWindow = null;
        _loaded.Clear();
        Window = IntPtr.Zero;
        PresenterWindow = IntPtr.Zero;
    }

    public async Task LoadAsync(string path, ScreenBounds bounds)
    {
        try { await OnSta(() => { Load(path, bounds); return true; }); }
        catch (Exception ex) { App.LogError(ex); }
    }

    private dynamic Load(string path, ScreenBounds bounds)
    {
        var app = Instance(bounds);
        if (_loaded.TryGetValue(path, out var known))
        {
            try { _ = ((dynamic)known).Slides.Count; return known; }
            catch { _loaded.Remove(path); } // closed meanwhile (by the operator, by PowerPoint)
        }
        var clock = Stopwatch.StartNew();
        // ReadOnly, not Untitled, WithWindow: false -> no editing window on the operator's screen.
        dynamic presentation = app.Presentations.Open(path, MsoTrue, MsoFalse, MsoFalse);
        _loaded[path] = presentation;
        Log($"loaded {Path.GetFileName(path)} in {clock.ElapsedMilliseconds} ms ({(int)presentation.Slides.Count} slides)");
        // At most the one on air plus the one just loaded: a long running order must not keep
        // dozens of files open in PowerPoint.
        foreach (var other in _loaded.Keys.ToList())
        {
            if (string.Equals(other, path, StringComparison.OrdinalIgnoreCase)) continue;
            dynamic loaded = _loaded[other];
            if (_presentation is not null && ReferenceEquals(loaded, _presentation)) continue;
            try { loaded.Close(); } catch { /* already closed */ }
            _loaded.Remove(other);
        }
        return presentation;
    }

    public async Task<PlayerState> OpenAsync(string path, ScreenBounds bounds, int startSlide, ScreenBounds? presenter)
    {
        try { return await OnSta(() => OpenCore(path, bounds, startSlide, presenter)); }
        catch (Exception ex)
        {
            // Most often a PowerPoint instance that went away under us: once more with a new one.
            Log($"show of {Path.GetFileName(path)} FAILED: {ex.GetType().Name}: {ex.Message} - trying again");
            return await OnSta(() =>
            {
                CloseCore(); // never leave a show of the failed try running underneath
                Forget();
                return OpenCore(path, bounds, startSlide, presenter);
            });
        }
    }

    private PlayerState OpenCore(string path, ScreenBounds bounds, int startSlide, ScreenBounds? presenter)
    {
        try { return OpenCoreUnsafe(path, bounds, startSlide, presenter); }
        catch
        {
            // A show that started but could not be taken over must not stay on screen under our
            // windows (it would also play its movies, unseen).
            CloseCore();
            throw;
        }
    }

    private PlayerState OpenCoreUnsafe(string path, ScreenBounds bounds, int startSlide, ScreenBounds? presenter)
    {
        CloseCore();
        _presentation = Load(path, bounds);
        var settings = _presentation.SlideShowSettings;
        settings.ShowType = PpShowTypeSpeaker;
        // PowerPoint's presenter view only when LiveSlides has a screen for it: PowerPoint would
        // put it on a screen of its own choosing, often the operator's.
        try { settings.ShowPresenterView = presenter is null ? MsoFalse : MsoTrue; }
        catch (Exception ex) { Log("ShowPresenterView not available: " + ex.Message); }
        // The show is started only now, when it goes on air (never earlier, never paused): its
        // animations, timings and movies start in front of the audience, once.
        var clock = Stopwatch.StartNew();
        _showWindow = settings.Run();
        Log($"show of {Path.GetFileName(path)} started in {clock.ElapsedMilliseconds} ms");

        Window = FindShowWindow(bounds);
        var pid = NativeWindows.ProcessIdOf(Window);
        if (!NativeWindows.Covers(Window, bounds))
        {
            NativeWindows.MoveTo(Window, bounds);
            Log($"show window moved onto the output screen {bounds.X},{bounds.Y} {bounds.Width}x{bounds.Height}");
        }

        PresenterWindow = IntPtr.Zero;
        if (presenter is { } presenterScreen)
        {
            // The presenter view: PowerPoint's other visible window (not the slide show, not the
            // editing window), whatever its title; it may appear a moment after the show.
            var show = Window;
            PresenterWindow = NativeWindows.MoveNewWindowAsync(pid, new[] { show },
                hwnd => NativeWindows.ClassName(hwnd) != EditingWindowClass,
                presenterScreen, 4000, requireTitle: false, minWidth: 400, minHeight: 300).GetAwaiter().GetResult();
            Log(PresenterWindow == IntPtr.Zero
                ? "presenter view NOT FOUND (the presenter screen shows the slide picture instead)"
                : $"presenter view 0x{PresenterWindow.ToInt64():X} ({NativeWindows.ClassName(PresenterWindow)}) moved onto the presenter screen");
        }

        var view = _showWindow.View;
        if (startSlide > 0) view.GotoSlide(startSlide + 1, MsoTrue);
        var state = ReadState();
        Log($"on air: slide {state.Slide + 1} of {state.SlideCount}, state {(int)view.State}");
        return state;
    }

    // Window classes of PowerPoint: the slide show, and the editing window (never moved).
    private const string ShowWindowClass = "screenClass";
    private const string EditingWindowClass = "PPTFrameClass";

    private static HashSet<int> PowerPointProcesses()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("POWERPNT"))
            using (process) ids.Add(process.Id);
        return ids;
    }

    /// <summary>
    /// The slide show window. Found among PowerPoint's windows, not asked to PowerPoint: on the
    /// first PC with PowerPoint, SlideShowWindow.HWND failed ("Error while invoking HWND") right
    /// after a successful start. It is PowerPoint's "screenClass" window (the one on the output
    /// screen if there are several); the HWND property is only a last resort.
    /// </summary>
    private IntPtr FindShowWindow(ScreenBounds bounds)
    {
        var clock = Stopwatch.StartNew();
        List<NativeWindows.WindowInfo> windows = new();
        while (clock.ElapsedMilliseconds < 5000)
        {
            windows = NativeWindows.VisibleWindows(PowerPointProcesses());
            var shows = windows.Where(w => w.Class == ShowWindowClass && w.Rect.Width >= 400 && w.Rect.Height >= 300).ToList();
            var chosen = shows.FirstOrDefault(w => NativeWindows.Covers(w.Handle, bounds))
                         ?? shows.OrderByDescending(w => (long)w.Rect.Width * w.Rect.Height).FirstOrDefault();
            if (chosen is not null)
            {
                Log($"slide show window found after {clock.ElapsedMilliseconds} ms: {chosen}{Environment.NewLine}  windows of PowerPoint:{Environment.NewLine}    {string.Join(Environment.NewLine + "    ", windows)}");
                return chosen.Handle;
            }
            Thread.Sleep(50);
        }

        Log($"no {ShowWindowClass} window after 5 s; windows of PowerPoint:{Environment.NewLine}    {string.Join(Environment.NewLine + "    ", windows)}");
        try
        {
            var handle = new IntPtr((int)_showWindow!.HWND);
            Log($"slide show window from SlideShowWindow.HWND: 0x{handle.ToInt64():X}");
            return handle;
        }
        catch (Exception ex) { Log("SlideShowWindow.HWND failed too: " + ex.Message); }

        // Last resort: PowerPoint's biggest window that is not the editing window.
        var biggest = windows.Where(w => w.Class != EditingWindowClass)
            .OrderByDescending(w => (long)w.Rect.Width * w.Rect.Height).FirstOrDefault();
        if (biggest is not null && biggest.Rect.Width >= 640)
        {
            Log($"slide show window taken to be the biggest one: {biggest}");
            return biggest.Handle;
        }
        throw new InvalidOperationException("finestra della presentazione di PowerPoint non trovata");
    }

    /// <summary>
    /// PowerPoint has no API to choose the slide show monitor; it reads this per-user option,
    /// the same one its "Monitor" drop-down writes. Set before PowerPoint starts.
    /// </summary>
    private static void PreferMonitor(ScreenBounds bounds)
    {
        try
        {
            var device = Displays.GetTargets().FirstOrDefault(t => t.Bounds == bounds)?.Id;
            if (device is null) return;
            foreach (var version in new[] { "16.0", "15.0" })
            {
                using var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{version}\PowerPoint\Options", true);
                if (key is null) continue;
                key.SetValue("DisplayMonitor", device, RegistryValueKind.String);
                key.SetValue("UseAutoMonSelection", 0, RegistryValueKind.DWord);
                Log($"slide show monitor set to {device} (Office {version})");
            }
        }
        catch (Exception ex) { App.LogError(ex); }
    }

    /// <summary>The show is never paused (see OpenCore); this only undoes a pause from elsewhere.</summary>
    public Task BeforeRevealAsync() => OnSta(() =>
    {
        try { if (_showWindow is not null && (int)_showWindow.View.State == PpSlideShowPaused) _showWindow.View.State = PpSlideShowRunning; }
        catch { /* show already over */ }
        return true;
    });

    public Task<PlayerState> NextAsync() => OnSta(() => { TryView("next", view => view.Next()); return ReadState(); });
    public Task<PlayerState> PreviousAsync() => OnSta(() => { TryView("previous", view => view.Previous()); return ReadState(); });
    public Task<PlayerState> GoToAsync(int slide) => OnSta(() => { TryView("go to", view => view.GotoSlide(slide + 1, MsoTrue)); return ReadState(); });
    public Task<PlayerState> GetStateAsync() => OnSta(ReadState);

    private void TryView(string what, Action<dynamic> action)
    {
        if (_showWindow is null) { Log(what + ": no show running"); return; }
        try { action(_showWindow.View); }
        catch (Exception ex) { Log($"{what} FAILED: {ex.Message}"); }
    }

    private PlayerState ReadState()
    {
        if (_showWindow is null || _presentation is null) return PlayerState.Stopped;
        try
        {
            var view = _showWindow.View;
            if ((int)view.State == PpSlideShowDone) return PlayerState.Stopped;
            int count = _presentation.Slides.Count;
            int position = view.CurrentShowPosition; // 1-based
            int clickIndex = 0, clickCount = 0;
            try { clickIndex = view.GetClickIndex(); clickCount = view.GetClickCount(); }
            catch { /* PowerPoint older than 2010: slide-level precision only */ }
            return new PlayerState(true, position - 1, count,
                AtStart: position <= 1 && clickIndex <= 0,
                AtEnd: position >= count && clickIndex >= clickCount);
        }
        catch (Exception ex)
        {
            Log("state not readable (show over?): " + ex.Message);
            return PlayerState.Stopped;
        }
    }

    public Task CloseAsync() => OnSta(() => { CloseCore(); Unload(); return true; });

    private void CloseCore()
    {
        Window = IntPtr.Zero;
        PresenterWindow = IntPtr.Zero;
        try { _showWindow?.View.Exit(); } catch { /* already closed */ }
        _showWindow = null;
        _presentation = null;
    }

    /// <summary>End of the show: every file loaded for it is closed.</summary>
    private void Unload()
    {
        foreach (dynamic presentation in _loaded.Values)
        {
            try { presentation.Close(); } catch { /* already closed */ }
        }
        _loaded.Clear();
    }

    public void Dispose()
    {
        if (_work.IsAddingCompleted) return;
        _work.Add(() =>
        {
            CloseCore();
            Unload();
            // Quit only a PowerPoint that has nothing else open: it may be the operator's own.
            try { if (_app is not null && (int)_app.Presentations.Count == 0) _app.Quit(); } catch { /* best effort */ }
            _app = null;
        });
        _work.CompleteAdding();
        _thread.Join(5000);
    }
}
