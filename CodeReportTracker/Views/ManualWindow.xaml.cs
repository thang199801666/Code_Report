using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;

namespace CodeReportTracker
{
    public partial class ManualWindow : Window
    {
        public ManualWindow()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Silently fail if unable to open browser
            }

            e.Handled = true;
        }
    }
}