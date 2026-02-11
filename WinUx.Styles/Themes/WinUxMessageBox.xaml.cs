using System;
using System.Windows;
using System.Windows.Input;
using WinUx.Controls;

namespace WinUx.Controls
{
    public partial class WinUxMessageBox : Window
    {
        public MessageBoxViewModel ViewModel { get; }

        public WinUxMessageBox()
        {
            InitializeComponent();
            ViewModel = new MessageBoxViewModel();
            this.DataContext = ViewModel;
        }

        public void Initialize(string message, MessageBoxButtons buttons, MessageBoxIcon icon, Action? onClose)
        {
            ViewModel.Setup(message, buttons, icon);
            // Ensure icon URI is sourced from the helper
            ViewModel.IconUri = MessageBoxHelper.GetIconUri(icon);
            ViewModel.RequestClose += onClose;
        }

        /// <summary>
        /// Show message box using automatic owner lookup (existing behavior).
        /// </summary>
        public static MessageBoxResults Show(string message, string title, MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.None)
        {
            // delegate to the overload that accepts an explicit parent (pass null)
            return Show((Window?)null, message, title, buttons, icon);
        }

        /// <summary>
        /// Show message box with an explicit parent window. Uses MessageBoxIcon only.
        /// </summary>
        public static MessageBoxResults Show(Window? parent, string message, string title, MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.None)
        {
            MessageBoxResults? result = null;

            // If caller supplied a parent, prefer it (only if it's usable).
            Window? ownerWindow = parent;
            if (ownerWindow == null)
            {
                var app = Application.Current;
                if (app != null)
                {
                    ownerWindow = app.Windows
                        .OfType<Window>()
                        .FirstOrDefault(w => w.IsLoaded && w.IsVisible);

                    // Fallback to MainWindow only if it's loaded and visible
                    if (ownerWindow == null && app.MainWindow != null && app.MainWindow.IsLoaded && app.MainWindow.IsVisible)
                        ownerWindow = app.MainWindow;
                }
            }
            else
            {
                // if a parent was supplied but not yet shown/visible, null it so we center on screen.
                if (!ownerWindow.IsLoaded || !ownerWindow.IsVisible)
                    ownerWindow = null;
            }

            var window = new WinUxMessageBox
            {
                Title = title
            };

            if (ownerWindow != null)
            {
                // Only set Owner when the chosen window has been shown previously.
                window.Owner = ownerWindow;
                window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
            {
                // No suitable owner -> center on screen
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            window.Initialize(message, buttons, icon, () =>
            {
                result = window.ViewModel.DialogResult;
                window.Close();
            });

            window.ShowDialog();
            return result ?? MessageBoxResults.None;
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                ViewModel.CancelCommand.Execute(null);
            else if (e.Key == Key.Enter)
                ViewModel.OkCommand.Execute(null);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(this);
        }
    }
}