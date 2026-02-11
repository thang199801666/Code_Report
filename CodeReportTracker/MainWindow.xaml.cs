using CodeReportTracker.Components.ViewModels;
using CodeReportTracker.Core.Models;
using CodeReportTracker.Core.Persistence;
using CodeReportTracker.Models;
using CodeReportTracker.ViewModels;
using Fluent;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;
using WinUx.Controls;
using WpfControls = System.Windows.Controls;

namespace CodeReportTracker
{
    public partial class MainWindow : RibbonWindow
    {
        private CancellationTokenSource? _tokenSource;
        private string? _currentFilePath;
        private readonly MainWindowViewModel _vm;
        private GridLength? _savedConsoleRowHeight;
        private bool _suppressExpanderEvents;
        private bool _isModified;
        private bool _suppressModifiedEvents;

        public MainWindow()
        {
            _suppressExpanderEvents = true;
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            try
            {
                _savedConsoleRowHeight = ConsoleRow.Height;
            }
            catch { /* Expected during initialization */ }

            _suppressExpanderEvents = false;
            _vm = CreateViewModel();
            DataContext = _vm;

            // Remove this line - don't load settings here
            // LoadAndApplyWindowSettings();
            AttachChangeTracking();
            ConfigureCommands();

            btnStop.IsEnabled = false;

            // Add this event handler for proper timing
            SourceInitialized += MainWindow_SourceInitialized;
        }

        #region Initialization
        private void MainWindow_SourceInitialized(object? sender, EventArgs e)
        {
            LoadAndApplyWindowSettings();
        }

        private MainWindowViewModel CreateViewModel()
        {
            return new MainWindowViewModel(
                searchAction: async () =>
                {
                    _tokenSource = new CancellationTokenSource();
                    Dispatcher.Invoke(() => SetRibbonButtons(false));
                    try
                    {
                        await Task.Run(() => _vm.UpdateDateTimeAsync(_tokenSource.Token));
                    }
                    finally
                    {
                        Dispatcher.Invoke(() => SetRibbonButtons(true));
                    }
                },
                stopAction: () =>
                {
                    _tokenSource?.Cancel();
                    PrintCommand("Cancel");
                },
                selectExcelAction: () => SelectExcelFile(this, null),
                exportAction: async () => await Task.Run(() => Dispatcher.Invoke(() => btnExport_Click(this, new RoutedEventArgs())))
            );
        }

        private void ConfigureCommands()
        {
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Save, SaveCommand_Executed, SaveCommand_CanExecute));
            CommandBindings.Add(new CommandBinding(ApplicationCommands.Help, HelpCommand_Executed));

            InputBindings.Add(new KeyBinding(ApplicationCommands.Save, Key.S, ModifierKeys.Control));
            InputBindings.Add(new KeyBinding(ApplicationCommands.Help, Key.F1, ModifierKeys.None));
        }

        #endregion

        #region Settings Management

        private void LoadAndApplyWindowSettings()
        {
            try
            {
                var settingsPath = GetSettingsPath();
                if (!File.Exists(settingsPath))
                {
                    return;
                }

                var json = File.ReadAllText(settingsPath);
                var loadedSettings = JsonSerializer.Deserialize<AppSettings>(json, GetJsonOptions());

                if (loadedSettings?.Window != null)
                {
                    ApplyWindowSettings(loadedSettings.Window);
                }
            }
            catch
            {
            }
        }

        private void ApplyWindowSettings(WindowSettings win)
        {
            // Apply loaded settings or use defaults
            Left = win.Left > 0 ? win.Left : 100;
            Top = win.Top > 0 ? win.Top : 100;
            Width = win.Width >= 800 ? win.Width : 1000;
            Height = win.Height >= 600 ? win.Height : 800;

            if (win.WindowState >= 0 && win.WindowState <= 2)
            {
                WindowState = win.WindowState switch
                {
                    1 => WindowState.Minimized,
                    2 => WindowState.Maximized,
                    _ => WindowState.Normal
                };
            }
        }

        private void SaveWindowSettings()
        {
            try
            {
                var settingsPath = GetSettingsPath();
                var appSettings = LoadOrCreateAppSettings(settingsPath);

                appSettings.Window = new WindowSettings
                {
                    Left = Left,
                    Top = Top,
                    Width = Width,
                    Height = Height,
                    WindowState = WindowState switch
                    {
                        WindowState.Minimized => 1,
                        WindowState.Maximized => 2,
                        _ => 0
                    }
                };

                var jsonOutput = JsonSerializer.Serialize(appSettings, GetJsonOptions(writeIndented: true));
                File.WriteAllText(settingsPath, jsonOutput);

                PrintCommand($"Window settings saved to: {settingsPath}");
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to save window settings: {ex.Message}");
            }
        }

        private AppSettings LoadOrCreateAppSettings(string settingsPath)
        {
            if (File.Exists(settingsPath))
            {
                var json = File.ReadAllText(settingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json, GetJsonOptions()) ?? new AppSettings();
            }

            return new AppSettings { WebSettings = _vm?.Settings?.ToList() ?? new List<SettingEntry>() };
        }

        private static string GetSettingsPath()
        {
            var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            return Path.Combine(baseDir, "settings.json");
        }

        private static JsonSerializerOptions GetJsonOptions(bool writeIndented = false)
        {
            return new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                AllowTrailingCommas = true,
                WriteIndented = writeIndented,
                Encoder = writeIndented ? System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping : null
            };
        }

        #endregion

        #region Change Tracking

        private void AttachChangeTracking()
        {
            try
            {
                _suppressModifiedEvents = true;

                if (_vm?.Tabs != null)
                {
                    _vm.Tabs.CollectionChanged += Tabs_CollectionChanged;
                    foreach (var tvm in _vm.Tabs)
                        HookTab(tvm);
                }

                _isModified = false;
            }
            finally
            {
                _suppressModifiedEvents = false;
            }
        }

        private void Tabs_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (var obj in e.NewItems.OfType<TabViewModel>())
                    HookTab(obj);
            }

            if (e.OldItems != null)
            {
                foreach (var obj in e.OldItems.OfType<TabViewModel>())
                    UnhookTab(obj);
            }

            MarkModified();
        }

        private void HookTab(TabViewModel tvm)
        {
            if (tvm == null) return;

            tvm.PropertyChanged += Tab_PropertyChanged;
            tvm.Items.CollectionChanged += TabItems_CollectionChanged;

            foreach (var item in tvm.Items.OfType<INotifyPropertyChanged>())
                item.PropertyChanged += CodeItem_PropertyChanged;
        }

        private void UnhookTab(TabViewModel tvm)
        {
            if (tvm == null) return;

            tvm.PropertyChanged -= Tab_PropertyChanged;
            tvm.Items.CollectionChanged -= TabItems_CollectionChanged;

            foreach (var item in tvm.Items.OfType<INotifyPropertyChanged>())
                item.PropertyChanged -= CodeItem_PropertyChanged;
        }

        private void Tab_PropertyChanged(object? sender, PropertyChangedEventArgs e) => MarkModified();

        private void TabItems_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (var inpc in e.NewItems.OfType<INotifyPropertyChanged>())
                    inpc.PropertyChanged += CodeItem_PropertyChanged;
            }

            if (e.OldItems != null)
            {
                foreach (var inpc in e.OldItems.OfType<INotifyPropertyChanged>())
                    inpc.PropertyChanged -= CodeItem_PropertyChanged;
            }

            MarkModified();
        }

        private void CodeItem_PropertyChanged(object? sender, PropertyChangedEventArgs e) => MarkModified();

        private void MarkModified()
        {
            if (!_suppressModifiedEvents)
                _isModified = true;
        }

        #endregion

        #region Command Handlers

        private void SaveCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
            e.Handled = true;
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var saveAs = string.Equals(e?.Parameter as string, "SaveAs", StringComparison.OrdinalIgnoreCase);
            PerformSave(saveAs);
            e.Handled = true;
        }

        private void HelpCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            btnHelp_Click(sender, e);
            e.Handled = true;
        }

        #endregion

        #region Event Handlers

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            tbConsole.AppendText("**************Initialize**************\n");

            // Log settings status after console is ready
            try
            {
                var settingsPath = GetSettingsPath();
                if (File.Exists(settingsPath))
                {
                    PrintCommand($"Window settings loaded from: {settingsPath}");
                }
                else
                {
                    PrintCommand("No settings file found. Using default window size and position.");
                }
            }
            catch (Exception ex)
            {
                PrintCommand($"Settings check failed: {ex.Message}");
            }
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = e.Uri.AbsoluteUri, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to open link: {e.Uri.AbsoluteUri} - {ex.Message}");
                WinUxMessageBox.Show(
                    $"Could not open link:\n{e.Uri.AbsoluteUri}\n\n{ex.Message}",
                    "Open Link Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            e.Handled = true;
        }

        private void Console_TextChanged(object sender, TextChangedEventArgs e)
        {
            ((WpfControls.TextBox)sender).ScrollToEnd();
        }

        private void ConsoleExpander_Collapsed(object sender, RoutedEventArgs e)
        {
            if (_suppressExpanderEvents) return;

            _savedConsoleRowHeight = ConsoleRow.Height;
            ConsoleRow.Height = new GridLength(34.0, GridUnitType.Pixel);
        }

        private void ConsoleExpander_Expanded(object sender, RoutedEventArgs e)
        {
            if (_suppressExpanderEvents) return;

            ConsoleRow.Height = _savedConsoleRowHeight ?? new GridLength(1, GridUnitType.Star);
            _savedConsoleRowHeight = null;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!HandleUnsavedChanges())
            {
                e.Cancel = true;
                return;
            }

            SaveWindowSettings();
            _tokenSource?.Cancel();
        }

        #endregion

        #region UI Operations

        private void SelectExcelFile(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Filter = "Excel File (*.xlsx)|*.xlsx",
                Title = "Select Excel File",
                Multiselect = false
            };

            if (fileDialog.ShowDialog() == true)
            {
                PrintCommand("Import Excel File: " + fileDialog.FileName);
                MarkModified();
            }
        }

        private void PrintCommand(string input)
        {
            var output = $"-[{DateTime.Now:hh:mm:ss MM-dd-yyyy}] : {input}\n";

            void Write()
            {
                tbConsole.AppendText(output);
                _vm?.AppendConsole(output.TrimEnd('\n'));
            }

            if (Dispatcher?.CheckAccess() ?? false)
                Write();
            else
                Dispatcher.Invoke(Write);
        }

        private void SetRibbonButtons(bool enabled)
        {
            void UpdateButtons()
            {
                btnUpdateSplit.IsEnabled = enabled;
                btnExport.IsEnabled = enabled;
                btnOpen.IsEnabled = enabled;
                btnSaveSplit.IsEnabled = enabled;
                btnStop.IsEnabled = !enabled;
                btnDownloadSplit.IsEnabled = enabled;
            }

            if (Dispatcher.CheckAccess())
                UpdateButtons();
            else
                Dispatcher.Invoke(UpdateButtons);
        }

        private static void CloseSplitDropDown(object? splitButton)
        {
            if (splitButton == null) return;

            try
            {
                var type = splitButton.GetType();
                var prop = type.GetProperty("IsDropDownOpen") ?? type.GetProperty("IsOpen");

                if (prop?.CanWrite == true)
                {
                    prop.SetValue(splitButton, false);
                    return;
                }

                if (splitButton is Visual visual)
                {
                    var popup = FindVisualChild<Popup>(visual);
                    if (popup != null)
                        popup.IsOpen = false;
                }
            }
            catch { /* Best effort attempt */ }
        }

        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typed)
                    return typed;

                var result = FindVisualChild<T>(child);
                if (result != null)
                    return result;
            }

            return null;
        }

        #endregion

        #region Button Click Handlers

        private async void btnCheckLink_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource = new CancellationTokenSource();
            SetRibbonButtons(false);
            try
            {
                await _vm.CheckLinkAsync(_tokenSource.Token);
                MarkModified();
            }
            finally
            {
                SetRibbonButtons(true);
                CloseSplitDropDown(btnUpdateSplit);
            }
        }

        private async void btWebCheck_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource = new CancellationTokenSource();
            
        }

        private async void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource = new CancellationTokenSource();
            SetRibbonButtons(false);
            try
            {
                await Task.Run(() => _vm.UpdateDateTimeAsync(_tokenSource.Token));
                MarkModified();
            }
            finally
            {
                SetRibbonButtons(true);
                CloseSplitDropDown(btnUpdateSplit);
            }
        }

        private async void btnUpdateLocal_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource = new CancellationTokenSource();
            SetRibbonButtons(false);
            try
            {
                await ExecuteLocalUpdate();
            }
            finally
            {
                SetRibbonButtons(true);
                CloseSplitDropDown(btnUpdateSplit);
            }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource?.Cancel();
            PrintCommand("Cancel");
        }
        private async void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_vm?.Tabs == null || _vm.Tabs.Count == 0)
            {
                PrintCommand("Export aborted: no tabs to export.");
                WinUxMessageBox.Show(
                    "No data available to export.",
                    "Export to Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Export to Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                FileName = $"CodeReport_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };

            if (dlg.ShowDialog() != true || string.IsNullOrWhiteSpace(dlg.FileName))
            {
                PrintCommand("Export canceled by user.");
                return;
            }

            SetRibbonButtons(false);
            try
            {
                await Task.Run(() => ExportToExcel(dlg.FileName));
                PrintCommand($"Successfully exported to: {dlg.FileName}");

                WinUxMessageBox.Show(
                    $"Data exported successfully to:\n{dlg.FileName}",
                    "Export Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                PrintCommand($"Export failed: {ex.Message}");
                WinUxMessageBox.Show(
                    $"Failed to export data:\n{ex.Message}",
                    "Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetRibbonButtons(true);
            }
        }

        private void ExportToExcel(string filePath)
        {
            var excelController = new ExcelControls.ExcelController();

            // Define columns to export - First column is Code Report No with hyperlink, Web Type removed
            var columnDefinitions = new[]
            {
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "Link",
                    DisplayName = "Code Report No",
                    IsHyperlink = true,
                    HasColor = true,
                    DisplayText = (item) => ((CodeItem)item).Number ?? string.Empty
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "ProductCategory",
                    DisplayName = "Product Category"
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "Description",
                    DisplayName = "Description"
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "ProductsListed",
                    DisplayName = "Products Listed",
                    CenterAlign = true
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "LatestCode",
                    DisplayName = "Latest Code",
                    HasColor = true,
                    CenterAlign = true
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "IssueDate",
                    DisplayName = "Issue/Rev Date",
                    HasColor = true
                },
                new ExcelControls.ColumnDefinition
                {
                    PropertyName = "ExpirationDate",
                    DisplayName = "Expiration Date",
                    HasColor = true
                }
            };

            // Prepare tab data
            var tabData = _vm.Tabs
                .Where(tab => tab?.Items != null && tab.Items.Count > 0)
                .ToDictionary(tab => tab.Header ?? "Sheet", tab => tab.Items.ToList());

            // Export with custom columns
            excelController.ExportWithCustomColumns(
                filePath,
                tabData,
                columnDefinitions,
                propertyGetter: GetPropertyValueForExport,
                colorProvider: GetCellColor,
                logger: PrintCommand
            );
        }

        private static object? GetPropertyValueForExport(CodeItem item, string propertyName)
        {
            if (item == null) return null;

            var property = item.GetType().GetProperty(propertyName);
            return property?.GetValue(item);
        }

        private static ClosedXML.Excel.XLColor? GetCellColor(CodeItem item, string propertyName)
        {
            // Code Report No column: color based on CodeExists
            if (propertyName == "Link")
            {
                return item.CodeExists
                    ? ClosedXML.Excel.XLColor.LightGreen
                    : ClosedXML.Excel.XLColor.LightCoral;
            }

            // Latest Code, Issue/Rev Date, and Expiration Date columns
            if (propertyName == "LatestCode" || propertyName == "IssueDate" || propertyName == "ExpirationDate")
            {
                var oldPropertyName = propertyName + "_Old";
                var oldValue = item.GetType().GetProperty(oldPropertyName)?.GetValue(item)?.ToString() ?? string.Empty;
                var newValue = item.GetType().GetProperty(propertyName)?.GetValue(item)?.ToString() ?? string.Empty;

                // ALWAYS return yellow if old value exists
                if (!string.IsNullOrWhiteSpace(oldValue))
                {
                    // Values are different - yellow
                    if (!string.Equals(oldValue.Trim(), newValue.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        return ClosedXML.Excel.XLColor.Yellow;
                    }
                }

                // No color if no old value or values are same
                return null;
            }

            return null;
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Open Code Report (.crp)",
                Filter = "Code Report (*.crp)|*.crp|All files (*.*)|*.*",
                Multiselect = false
            };

            if (dlg.ShowDialog() != true || string.IsNullOrWhiteSpace(dlg.FileName))
                return;

            try
            {
                LoadCrpFile(dlg.FileName);
                _currentFilePath = dlg.FileName;
                _isModified = false;
                PrintCommand($"Loaded CRP file: {dlg.FileName}");
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to load CRP file: {ex.Message}");
                WinUxMessageBox.Show(
                    $"Failed to load CRP file:\n{ex.Message}",
                    "Load Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            PerformSave(saveAs: false);
            CloseSplitDropDown(btnSaveSplit);
        }

        private void btnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            PerformSave(saveAs: true);
            CloseSplitDropDown(btnSaveSplit);
        }

        private void btnSettings_Click(object sender, RoutedEventArgs e)
        {
            if (_vm == null)
            {
                PrintCommand("Settings: ViewModel not available.");
                return;
            }

            try
            {
                var win = new SettingsWindow(_vm) { Owner = this };
                if (win.ShowDialog() != true)
                    PrintCommand("Settings changes canceled.");
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to open Settings: {ex.Message}");
            }
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var manualWindow = new ManualWindow
                {
                    Owner = this,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                };
                manualWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to open manual: {ex.Message}");
                WinUxMessageBox.Show(
                    $"Unable to open user manual:\n{ex.Message}",
                    "Manual Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private async void btnDownload_Click(object sender, RoutedEventArgs e)
        {
            _tokenSource = new CancellationTokenSource();
            SetRibbonButtons(false);
            try
            {
                var tvm = GetSelectedTabViewModel();
                var destBase = ResolveRootForTabPdfFolder(tvm);
                await _vm.DownloadPdfsAsync(destBase, _tokenSource.Token);
                MarkModified();
            }
            finally
            {
                SetRibbonButtons(true);
                _tokenSource = null;
                CloseSplitDropDown(btnDownloadSplit);
            }
        }

        private void btnDeletePdfs_Click(object sender, RoutedEventArgs e)
        {
            var sel = MainTabView?.SelectedItem;
            if (sel == null)
            {
                PrintCommand("Delete PDFs aborted: no tab selected.");
                return;
            }

            var tvm = GetSelectedTabViewModel();
            var pdfFolder = ResolvePdfFolderForTab(tvm, sel);

            if (string.IsNullOrWhiteSpace(pdfFolder) || !Directory.Exists(pdfFolder))
            {
                PrintCommand($"Delete PDFs aborted: folder not found '{pdfFolder ?? "(null)"}'.");
                ShowDeleteInfo($"No PDF folder found for the selected tab:\n{pdfFolder}");
                CloseSplitDropDown(btnDownloadSplit);
                return;
            }

            ExecuteDeletePdfs(tvm, pdfFolder);
        }

        #endregion

        #region File Operations

        private bool HandleUnsavedChanges()
        {
            if (!_isModified) return true;

            var result = WinUxMessageBox.Show(
                "There are unsaved changes. Do you want to save before closing?",
                "Unsaved Changes",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning);

            if (result == MessageBoxResults.Cancel)
                return false;

            if (result == MessageBoxResults.Yes)
                return PerformSave(saveAs: string.IsNullOrWhiteSpace(_currentFilePath));

            return true;
        }

        private bool PerformSave(bool saveAs)
        {
            if (MainTabView == null)
            {
                PrintCommand("Save aborted: MainTabView not available.");
                return false;
            }

            if (!saveAs && !string.IsNullOrWhiteSpace(_currentFilePath))
                return SaveToFile(_currentFilePath);

            return ShowSaveDialog();
        }

        private bool SaveToFile(string filePath)
        {
            try
            {
                var tabs = MainTabView.GetTabModels();
                BinaryCrpSerializer.Save(filePath, tabs);
                PrintCommand($"Saved CRP file: {filePath}");
                _isModified = false;
                return true;
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to save CRP file: {ex.Message}");
                WinUxMessageBox.Show(
                    $"Failed to save CRP file:\n{ex.Message}",
                    "Save Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private bool ShowSaveDialog()
        {
            var dlg = new SaveFileDialog
            {
                Title = "Save Code Report (.crp)",
                Filter = "Code Report (*.crp)|*.crp|All files (*.*)|*.*",
                DefaultExt = ".crp",
                AddExtension = true,
                FileName = "report.crp"
            };

            if (dlg.ShowDialog() != true || string.IsNullOrWhiteSpace(dlg.FileName))
                return false;

            if (SaveToFile(dlg.FileName))
            {
                _currentFilePath = dlg.FileName;
                return true;
            }

            return false;
        }

        private void LoadCrpFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            var tabs = BinaryCrpSerializer.Load(filePath);
            if (tabs == null) return;

            var tvmList = tabs.Select(tm => CreateTabViewModel(tm)).ToList();

            if (_vm?.Tabs != null)
            {
                LoadTabsIntoViewModel(tvmList);
            }
            else if (MainTabView != null)
            {
                LoadTabsIntoView(tvmList);
            }
        }

        private TabViewModel CreateTabViewModel(TabModel tm)
        {
            var tvm = new TabViewModel(tm.Header, null);
            foreach (var ci in tm.Items ?? Enumerable.Empty<CodeItem>())
                tvm.Items.Add(ci);

            TryInitializePdfFolder(tvm);
            TryLoadPdfFilesAsync(tvm);

            return tvm;
        }

        private void LoadTabsIntoViewModel(List<TabViewModel> tvmList)
        {
            try
            {
                _suppressModifiedEvents = true;

                _vm.Tabs.Clear();
                foreach (var tvm in tvmList)
                    _vm.Tabs.Add(tvm);

                _vm.SelectedTab = _vm.Tabs.FirstOrDefault();
                _isModified = false;
            }
            finally
            {
                _suppressModifiedEvents = false;
            }
        }

        private void LoadTabsIntoView(List<TabViewModel> tvmList)
        {
            var observable = new System.Collections.ObjectModel.ObservableCollection<object>(tvmList.Cast<object>());
            MainTabView.ItemsSource = observable;
            MainTabView.SelectedItem = observable.FirstOrDefault();

            _suppressModifiedEvents = true;
            _isModified = false;
            _suppressModifiedEvents = false;
        }

        private static void TryInitializePdfFolder(TabViewModel? tvm)
        {
            if (tvm == null) return;

            try
            {
                var mi = tvm.GetType().GetMethod("InitializePdfFolder", new[] { typeof(string) });
                if (mi != null)
                {
                    var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
                    mi.Invoke(tvm, new object?[] { baseDir });
                }
            }
            catch { /* Best effort attempt */ }
        }

        private static void TryLoadPdfFilesAsync(TabViewModel? tvm)
        {
            if (tvm == null) return;

            try
            {
                var mi = tvm.GetType().GetMethod("LoadPdfFilesAsync", new[] { typeof(string) });
                if (mi != null)
                {
                    var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
                    var result = mi.Invoke(tvm, new object?[] { baseDir });
                    if (result is Task t)
                    {
                        _ = t.ContinueWith(_ => { }, TaskScheduler.Default);
                    }
                }
            }
            catch { /* Best effort attempt */ }
        }

        #endregion

        #region PDF Operations

        private async Task ExecuteLocalUpdate()
        {
            var sel = MainTabView?.SelectedItem;
            if (sel == null)
            {
                PrintCommand("Local update aborted: no tab selected.");
                return;
            }

            var tvm = GetSelectedTabViewModel();
            var pdfFolder = ResolvePdfFolderForTab(tvm, sel);

            if (!ValidatePdfFolder(pdfFolder))
                return;

            TryLoadPdfFilesAsync(tvm);

            var pdfFiles = EnumeratePdfFiles(tvm, pdfFolder);
            if (!ValidatePdfFiles(pdfFiles, pdfFolder))
                return;

            LogPdfFiles(pdfFiles, pdfFolder);

            await Task.Run(() => _vm.UpdateDateTimeLocalAsync(_tokenSource.Token));
            MarkModified();
        }

        private bool ValidatePdfFolder(string? pdfFolder)
        {
            if (!string.IsNullOrWhiteSpace(pdfFolder) && Directory.Exists(pdfFolder))
                return true;

            PrintCommand($"Warning: PDF folder does not exist: '{pdfFolder ?? "(null)"}'");
            PrintCommand("Create the folder and place PDF files there, then try again.");

            WinUxMessageBox.Show(
                $"PDF folder not found:\n{pdfFolder}\n\nCreate this folder and place PDF files there before running Local Update.",
                "Folder Not Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return false;
        }

        private bool ValidatePdfFiles(string[] pdfFiles, string pdfFolder)
        {
            if (pdfFiles.Length > 0)
                return true;

            PrintCommand("Warning: No PDF files found in the folder.");
            WinUxMessageBox.Show(
                $"No PDF files found in:\n{pdfFolder}\n\nPlace PDF files in this folder before running Local Update.",
                "No PDFs Found",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return false;
        }

        private void LogPdfFiles(string[] pdfFiles, string pdfFolder)
        {
            PrintCommand($"Found {pdfFiles.Length} PDF file(s) in '{pdfFolder}'");

            foreach (var pdfFile in pdfFiles.Take(10))
            {
                PrintCommand($"  - {Path.GetFileName(pdfFile)}");
            }

            if (pdfFiles.Length > 10)
            {
                PrintCommand($"  ... and {pdfFiles.Length - 10} more");
            }
        }

        private void ExecuteDeletePdfs(TabViewModel? tvm, string pdfFolder)
        {
            string[] files;
            try
            {
                files = EnumeratePdfFiles(tvm, pdfFolder);
            }
            catch (Exception ex)
            {
                PrintCommand($"Failed to enumerate files in '{pdfFolder}': {ex.Message}");
                CloseSplitDropDown(btnDownloadSplit);
                return;
            }

            if (files.Length == 0)
            {
                PrintCommand($"No PDF files to delete in '{pdfFolder}'.");
                ShowDeleteInfo($"No PDF files found in:\n{pdfFolder}");
                CloseSplitDropDown(btnDownloadSplit);
                return;
            }

            if (!ConfirmDelete(files.Length, pdfFolder))
            {
                PrintCommand("Delete PDFs canceled by user.");
                CloseSplitDropDown(btnDownloadSplit);
                return;
            }

            var deleted = DeletePdfFiles(tvm, files, pdfFolder);
            ShowDeleteResult(deleted, pdfFolder);
            MarkModified();

            TryLoadPdfFilesAsync(tvm);
            CloseSplitDropDown(btnDownloadSplit);
        }

        private bool ConfirmDelete(int fileCount, string pdfFolder)
        {
            var result = WinUxMessageBox.Show(
                $"Delete {fileCount} PDF file(s) in:\n{pdfFolder}\n\nThis action cannot be undone. Continue?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            return result == MessageBoxResults.Yes;
        }

        private int DeletePdfFiles(TabViewModel? tvm, string[] files, string pdfFolder)
        {
            var toDelete = BuildDeleteList(tvm, files, pdfFolder);
            int deleted = 0;

            foreach (var path in toDelete)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        deleted++;
                    }
                }
                catch (Exception ex)
                {
                    PrintCommand($"Failed to delete '{path}': {ex.Message}");
                }
            }

            return deleted;
        }

        private void ShowDeleteResult(int deleted, string pdfFolder)
        {
            PrintCommand($"Deleted {deleted} PDF(s) from '{pdfFolder}'.");
            WinUxMessageBox.Show(
                $"Deleted {deleted} PDF(s) from:\n{pdfFolder}",
                "Delete PDFs",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ShowDeleteInfo(string message)
        {
            WinUxMessageBox.Show(message, "Delete PDFs", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region Helper Methods

        private TabViewModel? GetSelectedTabViewModel()
        {
            try
            {
                var sel = MainTabView?.SelectedItem;
                return sel switch
                {
                    TabViewModel tvm => tvm,
                    TabItem ti => ti.DataContext as TabViewModel,
                    _ => null
                };
            }
            catch
            {
                return null;
            }
        }

        private static string? ResolveRootForTabPdfFolder(TabViewModel? tvm)
        {
            if (tvm == null || string.IsNullOrWhiteSpace(tvm.PdfFolder))
                return null;

            try
            {
                var parent = Path.GetDirectoryName(tvm.PdfFolder);
                if (string.IsNullOrWhiteSpace(parent)) return null;

                var root = Path.GetDirectoryName(parent);
                return string.IsNullOrWhiteSpace(root) ? null : root;
            }
            catch
            {
                return null;
            }
        }

        private static string ResolvePdfFolderForTab(TabViewModel? tvm, object sel)
        {
            if (tvm != null && !string.IsNullOrWhiteSpace(tvm.PdfFolder))
                return tvm.PdfFolder;

            var tabName = sel switch
            {
                TabItem ti => ti.Header?.ToString() ?? "Unknown",
                TabViewModel tvm2 => tvm2.Header ?? "Unknown",
                _ => "Unknown"
            };

            var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            var safeTabName = Regex.Replace(tabName, @"\s+", "_");
            return Path.Combine(baseDir, "Pdf Files", safeTabName);
        }

        private static string[] EnumeratePdfFiles(TabViewModel? tvm, string pdfFolder)
        {
            if (tvm?.PdfFiles != null && tvm.PdfFiles.Count > 0)
            {
                var list = new List<string>(tvm.PdfFiles.Count);
                foreach (var entry in tvm.PdfFiles)
                {
                    if (string.IsNullOrWhiteSpace(entry)) continue;

                    try
                    {
                        var full = Path.IsPathRooted(entry)
                            ? Path.GetFullPath(entry)
                            : Path.GetFullPath(Path.Combine(pdfFolder, entry));

                        if (File.Exists(full))
                            list.Add(full);
                    }
                    catch { /* Skip invalid paths */ }
                }

                if (list.Count > 0)
                    return list.ToArray();
            }

            return Directory.GetFiles(pdfFolder, "*.pdf");
        }

        private static HashSet<string> BuildDeleteList(TabViewModel? tvm, IEnumerable<string> files, string pdfFolder)
        {
            var toDelete = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var filesArray = files.ToArray();

            if (tvm != null)
            {
                foreach (var code in tvm.Items)
                {
                    AddCandidateFilesToDelete(tvm, code, filesArray, pdfFolder, toDelete);
                }
            }

            if (toDelete.Count == 0)
            {
                foreach (var f in filesArray)
                    toDelete.Add(Path.GetFullPath(f));
            }

            return toDelete;
        }

        private static void AddCandidateFilesToDelete(
            TabViewModel tvm,
            CodeItem code,
            string[] filesArray,
            string pdfFolder,
            HashSet<string> toDelete)
        {
            try
            {
                var resolved = tvm.GetCandidatePdfPath(code);
                if (!string.IsNullOrWhiteSpace(resolved) && File.Exists(resolved))
                {
                    toDelete.Add(Path.GetFullPath(resolved));
                    return;
                }

                var suggested = GetSuggestedFileName(code);
                if (string.IsNullOrWhiteSpace(suggested)) return;

                var safe = MakeSafeFileNameLocal(suggested);
                if (!safe.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    safe += ".pdf";

                var candidateFull = Path.Combine(pdfFolder, safe);
                if (File.Exists(candidateFull))
                {
                    toDelete.Add(Path.GetFullPath(candidateFull));
                    return;
                }

                AddMatchingFiles(safe, filesArray, toDelete);
            }
            catch { /* Skip problematic items */ }
        }

        private static string? GetSuggestedFileName(CodeItem code)
        {
            return !string.IsNullOrWhiteSpace(code.Number) ? code.Number
                : !string.IsNullOrWhiteSpace(code.LatestCode) ? code.LatestCode
                : !string.IsNullOrWhiteSpace(code.Link) ? Path.GetFileName(code.Link)
                : null;
        }

        private static void AddMatchingFiles(string safe, string[] filesArray, HashSet<string> toDelete)
        {
            var baseNoExt = Path.GetFileNameWithoutExtension(safe);

            foreach (var f in filesArray)
            {
                var name = Path.GetFileName(f);
                if (string.Equals(name, safe, StringComparison.OrdinalIgnoreCase) ||
                    name.StartsWith(baseNoExt + "(", StringComparison.OrdinalIgnoreCase) ||
                    name.StartsWith(baseNoExt + " (", StringComparison.OrdinalIgnoreCase))
                {
                    toDelete.Add(Path.GetFullPath(f));
                }
            }
        }

        private static string MakeSafeFileNameLocal(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Unnamed";

            var invalid = Path.GetInvalidFileNameChars();
            var chars = input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
            var safe = new string(chars);

            safe = Regex.Replace(safe, @"\s{2,}", " ").Trim();
            safe = Regex.Replace(safe, @"[\. ]+$", "");

            return string.IsNullOrEmpty(safe) ? "Unnamed" : safe;
        }

        #endregion

        #region Nested Classes

        private class AppSettings
        {
            public List<SettingEntry>? WebSettings { get; set; }
            public WindowSettings? Window { get; set; }
        }

        private class WindowSettings
        {
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public int WindowState { get; set; }
        }

        #endregion
    }
}