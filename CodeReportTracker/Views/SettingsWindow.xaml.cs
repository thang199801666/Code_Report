using CodeReportTracker.Core.Models;
using CodeReportTracker.Models;
using CodeReportTracker.ViewModels;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Navigation;
using WinUx.Controls;

namespace CodeReportTracker
{
    public partial class SettingsWindow : Window
    {
        private readonly MainWindowViewModel _vm;
        private bool _cpuWarningShown;

        public SettingsWindow(MainWindowViewModel vm)
        {
            InitializeComponent();
            _vm = vm ?? throw new ArgumentNullException(nameof(vm));
            DataContext = _vm;
            _vm.RefreshAvailableModels();
            _cpuWarningShown = _vm.CpuPercent > 75;
        }

        private void OpenModelFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dir = _vm.AiModelFolderPath;
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                Process.Start(new ProcessStartInfo
                {
                    FileName = dir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                WinUxMessageBox.Show("Failed to open AI models folder: " + ex.Message,
                    "Open Folder Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CpuPercentSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_vm == null) return;

            if (e.NewValue > 75 && !_cpuWarningShown)
            {
                _cpuWarningShown = true;
                WinUxMessageBox.Show(
                    $"Using more than 75% of the available CPUs ({_vm.CpuPercent}%) for the AI model may slow down other applications.\n\nYou can still continue.",
                    "CPU Usage Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else if (e.NewValue <= 75)
            {
                _cpuWarningShown = false;
            }
        }

        // Open HTTP/HTTPS links in default browser
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                var uri = e.Uri;
                if (uri == null || string.IsNullOrWhiteSpace(uri.AbsoluteUri))
                {
                    return;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = uri.AbsoluteUri,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                WinUxMessageBox.Show("Failed to open link: " + ex.Message, 
                    "Open Link Error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }

            e.Handled = true;
        }

        // Open local PDF folder (or path) in Explorer. Hyperlink Click is used so we get DataContext.
        private void PdfFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is System.Windows.Documents.Hyperlink hl && hl.DataContext is SettingEntry entry)
                {
                    var path = entry?.PdfFolder;
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        return;
                    }

                    // If path looks like a file:// URI, use LocalPath; otherwise pass path directly.
                    string target = path;
                    if (Uri.TryCreate(path, UriKind.Absolute, out var maybeUri) && maybeUri.IsFile)
                        target = maybeUri.LocalPath;

                    if (Directory.Exists(target) || File.Exists(target))
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = target,
                            UseShellExecute = true
                        };
                        Process.Start(psi);
                    }
                    else
                    {
                        WinUxMessageBox.Show("The specified folder or file does not exist:\n" + target, 
                            "Not Found", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                WinUxMessageBox.Show("Failed to open folder: " + ex.Message, 
                    "Open Folder Error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Save all settings to settings.json in executable directory
                _vm.SaveSettings();

                WinUxMessageBox.Show("Settings saved successfully!", 
                    "Settings", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                WinUxMessageBox.Show("Failed to save settings: " + ex.Message, 
                                     "Save Error", 
                                     MessageBoxButtons.OK, 
                                     MessageBoxIcon.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}