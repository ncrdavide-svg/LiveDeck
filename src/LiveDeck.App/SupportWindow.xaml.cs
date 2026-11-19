using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace LiveDeck.App;

/// <summary>
/// A small modal "support the app" window shown once per launch (see MainWindow's Loaded
/// handler), as in LiveCue and LiveCount: a Buy Me a Coffee QR code/link, styled in BMC's own brand colors. "No, grazie"
/// stays disabled for a few seconds - long enough to actually notice the QR code - rather than
/// being instantly dismissible.
/// </summary>
public partial class SupportWindow : Window
{
    private const string SupportUrl = "https://buymeacoffee.com/cin8";
    private const int WaitSeconds = 3;

    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromSeconds(1) };
    private int _secondsLeft = WaitSeconds;

    public SupportWindow()
    {
        InitializeComponent();
        UpdateButtonText();
        _timer.Tick += (_, _) =>
        {
            _secondsLeft--;
            if (_secondsLeft <= 0)
            {
                _timer.Stop();
                NoThanksButton.IsEnabled = true;
                NoThanksButton.Content = "No, grazie";
            }
            else
            {
                UpdateButtonText();
            }
        };
        _timer.Start();
    }

    private void UpdateButtonText() => NoThanksButton.Content = $"No, grazie ({_secondsLeft})";

    private void NoThanksButton_Click(object sender, RoutedEventArgs e) => Close();

    private void Link_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(SupportUrl) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            App.LogError(ex);
        }
    }
}
