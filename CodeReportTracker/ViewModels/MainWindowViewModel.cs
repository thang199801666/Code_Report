using CodeReportTracker.Components.ViewModels;
using CodeReportTracker.Core.Models;
using CodeReportTracker.Models;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using PDFControls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace CodeReportTracker.ViewModels
{
    public partial class MainWindowViewModel : ObservableObject
    {
        #region Constants

        private const int DefaultHttpTimeoutSeconds = 10;
        private const int ExtendedHttpTimeoutSeconds = 60;
        private const int CpuLimitPercent = 75;
        private const int SniffBufferSize = 8192;
        private const int FileBufferSize = 81920;
        private const string DefaultFileName = "Unnamed";
        private const string PdfExtension = ".pdf";
        private const string PdfFolderName = "Pdf Files";
        private const string SettingsFileName = "settings.json";
        private const string DefaultLatestCodeContexts = "following code";
        private const string DefaultIssueDateContexts = "issue;revised;date of revision";
        private const string DefaultExpirationDateContexts = "renewal;renewal date;renew;valid through;valid thru;valid until;active through;active until;available until;ends on;expiration;expires;expiration date;through the end of;through end of;compliance program through;program through";
        private const string DefaultAiModelFileName = "SmolLM2-135M-Instruct-Q4_K_M.gguf";
        private const string ProductCountCacheFileName = "products_cache.json";
        private const int ConsoleMaxLength = 100000;

        #endregion

        #region Fields

        private string _consoleText = string.Empty;
        private readonly List<ConsoleLine> _consoleBuffer = new List<ConsoleLine>();
        private readonly object _consoleLock = new object();
        private readonly DispatcherTimer _consoleTimer;

        public event Action<ConsoleLine>? ConsoleLineAdded;
        private bool _isBusy;
        private TabViewModel? _selectedTab;
        private readonly string _currentDir = Directory.GetCurrentDirectory();
        private AiPdfTextExtractor? _aiExtractor;
        private string? _aiModelFileNameUsed;
        private readonly Dictionary<string, int> _productCountCache = new(StringComparer.OrdinalIgnoreCase);
        private double _updateProgress;
        private bool _isUpdatingProgress;
        private int _updateProcessed;
        private int _updateTotal;

        // Delegates provided by the View for view-only operations
        private readonly Func<Task>? _searchAction;
        private readonly Action? _stopAction;
        private readonly Action? _selectExcelAction;
        private readonly Func<Task>? _exportAction;

        #endregion

        #region Properties

        public ObservableCollection<SettingEntry> Settings { get; }
        public ObservableCollection<ExtractionContextSettings> ContextSettings { get; }
        public ObservableCollection<TabViewModel> Tabs { get; }

        public AiMode ExtractionMode { get; set; } = AiMode.Ai;
        public bool CalculateProduct { get; set; } = false;
        public int CpuPercent { get; set; } = 75;
        public string AiModelFileName { get; set; } = DefaultAiModelFileName;

        public ObservableCollection<string> AvailableModels { get; } = new ObservableCollection<string>();
        public string AiModelFolderPath => Path.Combine(AppContext.BaseDirectory, "AIModels");

        public TabViewModel? SelectedTab
        {
            get => _selectedTab;
            set => SetProperty(ref _selectedTab, value);
        }

        public string ConsoleText
        {
            get => _consoleText;
            set => SetProperty(ref _consoleText, value);
        }

        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                if (SetProperty(ref _isBusy, value))
                {
                    SearchCommand.NotifyCanExecuteChanged();
                    StopCommand.NotifyCanExecuteChanged();
                    SelectExcelCommand.NotifyCanExecuteChanged();
                    ExportCommand.NotifyCanExecuteChanged();
                    CloseTabCommand.NotifyCanExecuteChanged();
                }
            }
        }

        public double UpdateProgress
        {
            get => _updateProgress;
            private set => SetProperty(ref _updateProgress, value);
        }

        public bool IsUpdatingProgress
        {
            get => _isUpdatingProgress;
            private set => SetProperty(ref _isUpdatingProgress, value);
        }

        #endregion

        #region Constructor

        public MainWindowViewModel( Func<Task>? searchAction = null,
                                    Action? stopAction = null,
                                    Action? selectExcelAction = null,
                                    Func<Task>? exportAction = null)
        {
            _searchAction = searchAction;
            _stopAction = stopAction;
            _selectExcelAction = selectExcelAction;
            _exportAction = exportAction;

            ContextSettings = new ObservableCollection<ExtractionContextSettings>
            {
                new ExtractionContextSettings
                {
                    LatestCode = DefaultLatestCodeContexts,
                    IssueDate = DefaultIssueDateContexts,
                    ExpirationDate = DefaultExpirationDateContexts
                }
            };

            Settings = new ObservableCollection<SettingEntry>
            {
                new SettingEntry
                {
                    Name = "IAPMO",
                    Type = "ER",
                    Link = "https://forms.iapmo.org/ues_reports/EvaluationReports.aspx",
                    PdfFolder = "https://forms.iapmo.org/ues_reports/reports/"
                },
                new SettingEntry
                {
                    Name = "ICC-ES",
                    Type = "ESR",
                    Link = "https://icc-es.org/evaluation-report-program/reports-directory/",
                    PdfFolder = "https://cdn-v2.icc-es.org/wp-content/uploads/report-directory/"
                },
                new SettingEntry
                {
                    Name = "LADBS RR",
                    Type = "Other",
                    Link = "https://www.drjcertification.org/ter-directory",
                    PdfFolder = string.Empty
                }
            };

            Tabs = new ObservableCollection<TabViewModel> { new TabViewModel("New Tab") };
            SelectedTab = Tabs.FirstOrDefault();

            // Batch console output onto the UI thread so high-frequency logging does not
            // force WPF to re-render the whole console TextBox on every single line.
            _consoleTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
            _consoleTimer.Tick += (_, __) => FlushConsoleBuffer();
            _consoleTimer.Start();

            LoadProductCountCache();
        }

        private void LoadProductCountCache()
        {
            try
            {
                var path = Path.Combine(AppContext.BaseDirectory, ProductCountCacheFileName);
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var data = JsonSerializer.Deserialize<Dictionary<string, int>>(json, GetJsonOptions());
                    if (data != null)
                    {
                        _productCountCache.Clear();
                        foreach (var kv in data)
                            _productCountCache[kv.Key] = kv.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to load product count cache: {ex.Message}");
            }
        }

        private void SaveProductCountCache()
        {
            try
            {
                var path = Path.Combine(AppContext.BaseDirectory, ProductCountCacheFileName);
                var json = JsonSerializer.Serialize(_productCountCache, GetJsonOptions(writeIndented: true));
                File.WriteAllText(path, json);
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to save product count cache: {ex.Message}");
            }
        }

        private void FlushConsoleBuffer()
        {
            List<ConsoleLine> pending;
            lock (_consoleLock)
            {
                if (_consoleBuffer.Count == 0)
                    return;
                pending = new List<ConsoleLine>(_consoleBuffer);
                _consoleBuffer.Clear();
            }

            var sb = new StringBuilder();
            foreach (var line in pending)
            {
                ConsoleLineAdded?.Invoke(line);
                sb.Append(line.Text).Append(Environment.NewLine);
            }

            var next = ConsoleText + sb.ToString();
            if (next.Length > ConsoleMaxLength)
                next = next.Substring(next.Length - ConsoleMaxLength);
            ConsoleText = next;
        }

        #endregion

        #region Public Methods

        public void AppendConsole(string line) => AppendConsoleLine(line, ClassifyConsoleLevel(line));

        public void AppendInfo(string line) => AppendConsoleLine(line, ConsoleLevel.Info);

        public void AppendSuccess(string line) => AppendConsoleLine(line, ConsoleLevel.Success);

        public void AppendWarning(string line) => AppendConsoleLine(line, ConsoleLevel.Warning);

        public void AppendError(string line) => AppendConsoleLine(line, ConsoleLevel.Error);

        private void AppendConsoleLine(string line, ConsoleLevel level)
        {
            lock (_consoleLock)
                _consoleBuffer.Add(new ConsoleLine(line ?? string.Empty, level));
        }

        private static ConsoleLevel ClassifyConsoleLevel(string line)
        {
            var text = line ?? string.Empty;
            if (text.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("failed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("exception", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("unable to", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("cannot", StringComparison.OrdinalIgnoreCase) >= 0)
                return ConsoleLevel.Error;

            if (text.IndexOf("warning", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("cancelled", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("timed out", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("timeout", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("skipped", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("skipping", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("aborted", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("no link", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("no text", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("no rows", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("no matching", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("unavailable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("empty", StringComparison.OrdinalIgnoreCase) >= 0)
                return ConsoleLevel.Warning;

            if (text.IndexOf("successfully", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("finished.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("saved to", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("imported", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("downloaded", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("completed", StringComparison.OrdinalIgnoreCase) >= 0)
                return ConsoleLevel.Success;

            return ConsoleLevel.Info;
        }

        public void SaveSettings()
        {
            try
            {
                var settingsPath = GetSettingsPath();
                var appSettings = new AppSettings
                {
                    WebSettings = new List<SettingEntry>(Settings),
                    ContextSettings = ContextSettings.FirstOrDefault(),
                    ExtractionMode = ExtractionMode.ToString(),
                    CalculateProduct = CalculateProduct,
                    CpuPercent = CpuPercent,
                    AiModelFileName = AiModelFileName,
                };

                var json = JsonSerializer.Serialize(appSettings, GetJsonOptions(writeIndented: true));
                File.WriteAllText(settingsPath, json);

                AppendConsole($"Settings saved to: {settingsPath}");
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to save settings: {ex.Message}");
                throw;
            }
        }

        public void LoadSettings()
        {
            try
            {
                var settingsPath = GetSettingsPath();

                if (!File.Exists(settingsPath))
                {
                    AppendConsole("No settings file found. Using default settings.");
                    return;
                }

                var json = File.ReadAllText(settingsPath);
                var loadedSettings = JsonSerializer.Deserialize<AppSettings>(json, GetJsonOptions());

                if (loadedSettings?.WebSettings != null && loadedSettings.WebSettings.Count > 0)
                {
                    Settings.Clear();
                    foreach (var setting in loadedSettings.WebSettings)
                    {
                        Settings.Add(setting);
                    }
                }

                if (loadedSettings != null)
                {
                    ExtractionMode = ParseExtractionMode(loadedSettings.ExtractionMode);
                    CalculateProduct = loadedSettings.CalculateProduct ?? false;
                    CpuPercent = loadedSettings.CpuPercent ?? 75;
                    AiModelFileName = string.IsNullOrWhiteSpace(loadedSettings.AiModelFileName)
                        ? DefaultAiModelFileName
                        : loadedSettings.AiModelFileName;
                }

                if (loadedSettings?.ContextSettings != null)
                {
                    ContextSettings.Clear();
                    ContextSettings.Add(loadedSettings.ContextSettings);
                }
                EnsureContextDefaults();

                AppendConsole($"Settings loaded from: {settingsPath}");
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to load settings: {ex.Message}");
                AppendConsole("Using default settings.");
            }

            RefreshAvailableModels();
        }

        public void RefreshAvailableModels()
        {
            AvailableModels.Clear();
            try
            {
                var dir = AiModelFolderPath;
                if (Directory.Exists(dir))
                {
                    foreach (var file in Directory.EnumerateFiles(dir, "*.gguf", SearchOption.TopDirectoryOnly)
                        .OrderBy(f => Path.GetFileName(f), StringComparer.OrdinalIgnoreCase))
                    {
                        AvailableModels.Add(Path.GetFileName(file));
                    }
                }
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to scan AI models folder: {ex.Message}");
            }

            if (AvailableModels.Count == 0)
                AvailableModels.Add(DefaultAiModelFileName);

            if (!AvailableModels.Contains(AiModelFileName))
                AiModelFileName = AvailableModels[0];
        }

        #endregion

        #region Commands

        [RelayCommand(CanExecute = nameof(CanSearch))]
        private async Task SearchAsync()
        {
            if (_searchAction == null) return;
            IsBusy = true;
            try
            {
                await _searchAction();
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool CanSearch() => !IsBusy;

        [RelayCommand(CanExecute = nameof(CanStop))]
        private void Stop() => _stopAction?.Invoke();

        private bool CanStop() => IsBusy;

        [RelayCommand(CanExecute = nameof(CanSelectExcel))]
        private void SelectExcel() => _selectExcelAction?.Invoke();

        private bool CanSelectExcel() => !IsBusy;

        [RelayCommand(CanExecute = nameof(CanExport))]
        private async Task ExportAsync()
        {
            if (_exportAction == null) return;
            IsBusy = true;
            try
            {
                await _exportAction();
            }
            finally
            {
                IsBusy = false;
            }
        }

        private bool CanExport() => !IsBusy;

        [RelayCommand]
        private void CloseTab(object? item)
        {
            if (item is not TabViewModel tvm || !Tabs.Contains(tvm))
                return;

            var wasSelected = ReferenceEquals(SelectedTab, tvm);
            Tabs.Remove(tvm);

            if (wasSelected)
                SelectedTab = Tabs.FirstOrDefault();
        }

        #endregion

        #region Long-Running Operations

        public async Task CheckLinkAsync(CancellationToken token = default)
        {
            if (!ValidateSelectedTab("Check Link for Reports", out var selTab))
                return;

            var rowList = selTab.Items?.ToList() ?? new List<CodeItem>();
            if (!rowList.Any())
            {
                AppendConsole("Check Link for Reports: no rows to check.");
                return;
            }

            AppendConsole("Check Link for Reports started: checking PDF existence for selected table...");
            ResetPdfTimeoutState(rowList);

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(DefaultHttpTimeoutSeconds) };
            int maxConcurrency = CalculateMaxConcurrency(2);
            using var semaphore = new SemaphoreSlim(maxConcurrency);

            var settingsList = Settings?.ToList() ?? new List<SettingEntry>();

            var checkTasks = rowList.Select(code => CheckSingleCodeAsync(code, http, settingsList, semaphore, token)).ToList();
            await Task.WhenAll(checkTasks).ConfigureAwait(false);

            AppendConsole($"Check Link for Reports finished on tab '{SelectedTab?.Header ?? "Unknown"}'.");
        }


        public async Task UpdateDateTimeAsync(CancellationToken token = default)
        {
            if (!ValidateSelectedTab("Search", out var selTab))
                return;

            var rowList = selTab.Items?.ToList() ?? new List<CodeItem>();
            if (!rowList.Any())
            {
                AppendConsole("Search aborted: no rows to process.");
                return;
            }

            AppendConsole("Search started: reading first page of PDFs for selected table...");
            ResetPdfTimeoutState(rowList);
            StartProgress(rowList.Count);

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(ExtendedHttpTimeoutSeconds) };

            var maxConcurrency = CalculateMaxConcurrency(2);
            using var extractionSemaphore = new SemaphoreSlim(maxConcurrency);
            var extractedTextQueue = Channel.CreateBounded<(CodeItem code, string text)>(
                new BoundedChannelOptions(Math.Max(4, maxConcurrency * 2))
                {
                    FullMode = BoundedChannelFullMode.Wait,
                    SingleReader = true,
                    SingleWriter = false
                });

            // PDF extraction is parallel; AI consumption remains strictly single-threaded.
            var aiConsumer = ConsumeExtractedTextAsync(extractedTextQueue.Reader, token);
            var extractionTasks = rowList.Select(code =>
                ExtractCodeTextToQueueAsync(code, http, extractionSemaphore, extractedTextQueue.Writer, token)).ToList();

            try
            {
                await Task.WhenAll(extractionTasks).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                // cancelled - fall through so the consumer is awaited/stopped below.
            }
            finally
            {
                extractedTextQueue.Writer.TryComplete();
            }

            try
            {
                // The consumer stops as soon as the token is cancelled (see ConsumeExtractedTextAsync).
                await aiConsumer.ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                // cancelled - ignore.
            }

            EndProgress();
            AppendConsole(token.IsCancellationRequested ? "Search cancelled." : "Search finished.");
        }

        private void StartProgress(int total)
        {
            _updateProcessed = 0;
            _updateTotal = total;
            UpdateProgress = 0;
            IsUpdatingProgress = true;
        }

        private void EndProgress()
        {
            UpdateProgress = 0;
            IsUpdatingProgress = false;
        }

        private void AdvanceProgress()
        {
            var processed = Interlocked.Increment(ref _updateProcessed);
            var total = _updateTotal;
            UpdateProgress = total > 0 ? Math.Min(100.0, processed * 100.0 / total) : 0;
        }

        public async Task UpdateDateTimeLocalAsync(CancellationToken token = default)
        {
            if (!ValidateSelectedTab("Local update", out var selTab))
                return;

            var pdfFolderForTab = GetPdfFolderForTab(selTab);
            AppendConsole($"Local update: searching for PDFs in '{pdfFolderForTab}'");

            if (!Directory.Exists(pdfFolderForTab))
            {
                AppendConsole($"Warning: PDF folder does not exist: '{pdfFolderForTab}'");
                AppendConsole("No local PDFs found. Create the folder and place PDF files there.");
                return;
            }

            var rowList = selTab.Items?.ToList() ?? new List<CodeItem>();
            if (!rowList.Any())
            {
                AppendConsole("Local update aborted: no rows to process.");
                return;
            }

            AppendConsole("Local update started: reading first page of local PDFs for selected table...");

            int maxConcurrency = CalculateMaxConcurrency(2);
            using var semaphore = new SemaphoreSlim(maxConcurrency);

            var processTasks = rowList.Select(code => ProcessLocalCodeUpdateAsync(code, pdfFolderForTab, semaphore, token)).ToList();
            await Task.WhenAll(processTasks).ConfigureAwait(false);

            AppendConsole(token.IsCancellationRequested ? "Local update cancelled." : "Local update finished.");
        }

        public async Task DownloadPdfsAsync(string? destBase = null, CancellationToken token = default)
        {
            if (!ValidateSelectedTab("Download", out var selTab))
                return;

            var rowList = selTab.Items?.ToList() ?? new List<CodeItem>();
            if (!rowList.Any())
            {
                AppendConsole("Download aborted: no rows to download.");
                return;
            }

            var destFolder = PrepareDestinationFolder(selTab, destBase);
            if (destFolder == null)
                return;

            AppendConsole($"Download started: saving PDFs to {destFolder}");

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(ExtendedHttpTimeoutSeconds) };
            int maxConcurrency = CalculateMaxConcurrency(4);
            using var semaphore = new SemaphoreSlim(maxConcurrency);

            ResetDownloadProgress(rowList);

            var tasks = rowList.Select(code => DownloadSinglePdfAsync(code, destFolder, http, semaphore, token)).ToList();
            (CodeItem code, bool success, string message)[] results;
            try
            {
                results = await Task.WhenAll(tasks).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                AppendConsole("Download cancelled.");
                return;
            }

            ProcessDownloadResults(results);
            AppendConsole(token.IsCancellationRequested ? "Download cancelled." : "Download finished.");
        }

        #endregion

        #region Helper Methods - Validation

        private bool ValidateSelectedTab(string operationName, out TabViewModel selectedTab)
        {
            selectedTab = SelectedTab!;
            if (SelectedTab != null)
                return true;

            AppendConsole($"{operationName} aborted: no tab selected.");
            return false;
        }

        #endregion

        #region Helper Methods - Settings

        private static string GetSettingsPath()
        {
            var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            return Path.Combine(baseDir, SettingsFileName);
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

        #region Helper Methods - PDF Checking

        private async Task CheckSingleCodeAsync(
            CodeItem code,
            HttpClient http,
            List<SettingEntry> settingsList,
            SemaphoreSlim semaphore,
            CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return;

            var semaphoreEntered = false;
            try
            {
                await semaphore.WaitAsync(token).ConfigureAwait(false);
                semaphoreEntered = true;

                var (exists, newLink) = await CheckCodeExistsAsync(code, http, settingsList, token).ConfigureAwait(false);

                SafeDispatch(() =>
                {
                    if (!string.IsNullOrWhiteSpace(newLink))
                    {
                        code.Link = newLink;
                        AppendConsole($"Updated link for {code.Number} -> {newLink}");
                    }

                    code.CodeExists = exists;
                    code.HasCheck = true;
                    code.LastCheck = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
                });

                AppendConsole($"Checked {code.Number}: PDF {(exists ? "found" : "missing")}");
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                MarkPdfCheckTimedOut(code);
                AppendConsole($"Skipped {code.Number}: PDF URL timed out.");
            }
            catch (Exception ex)
            {
                AppendConsole($"Error checking {code.Number}: {ex.Message}");
            }
            finally
            {
                if (semaphoreEntered)
                    semaphore.Release();
            }
        }

        private async Task<(bool exists, string? newLink)> CheckCodeExistsAsync(
            CodeItem code,
            HttpClient http,
            List<SettingEntry> settingsList,
            CancellationToken token)
        {
            // Check local files first
            if (CheckLocalFilePaths(code, _currentDir))
                return (true, null);

            // Check existing link
            if (!string.IsNullOrWhiteSpace(code.Link) && Uri.IsWellFormedUriString(code.Link, UriKind.Absolute))
            {
                var uri = new Uri(code.Link);
                if ((uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) &&
                    await TryUrlHasPdfAsync(http, uri, token).ConfigureAwait(false))
                {
                    return (true, null);
                }

            }

            // Try mapping
            if (!string.IsNullOrWhiteSpace(code.Number))
            {
                var mappingResult = await TryFindMappedPdfAsync(http, code.Number, settingsList, token).ConfigureAwait(false);
                if (mappingResult.found)
                {
                    if (!string.IsNullOrWhiteSpace(mappingResult.foundLink))
                        AppendConsole($"Found PDF via mapping: {mappingResult.settingName} -> {mappingResult.foundLink}");
                    return (true, mappingResult.foundLink);
                }

                // ICC-ES fallback
                var matchedSetting = GetMatchingSettingForCode(code.Number);
                if (IsIccSource(code, matchedSetting))
                {
                    var fallbackLink = await TryIccFallbackAsync(http, code.Number, token).ConfigureAwait(false);
                    if (fallbackLink != null)
                        return (true, fallbackLink);
                }
            }

            return (false, null);
        }

        private async Task<string?> TryIccFallbackAsync(HttpClient http, string number, CancellationToken token)
        {
            try
            {
                var encoded = Uri.EscapeDataString(number.Trim());
                var fallback = $"https://icc-es.org/wp-content/uploads/report-directory/{encoded}.pdf";
                AppendConsole($"ICC-ES fallback check for {number}: {fallback}");

                if (await TryHeadIsPdfAsync(http, new Uri(fallback), token).ConfigureAwait(false))
                    return fallback;
            }
            catch { /* Best effort attempt */ }

            return null;
        }

        private static bool CheckLocalFilePaths(CodeItem code, string currentDir)
        {
            try
            {
                // Check URI file path
                if (!string.IsNullOrWhiteSpace(code.Link) && Uri.IsWellFormedUriString(code.Link, UriKind.Absolute))
                {
                    var uri = new Uri(code.Link);
                    if (uri.IsFile && File.Exists(uri.LocalPath))
                        return true;
                }

                // Check direct file path
                if (!string.IsNullOrWhiteSpace(code.Link) && !Uri.IsWellFormedUriString(code.Link, UriKind.Absolute) && File.Exists(code.Link))
                    return true;

                // Check number-based path
                if (!string.IsNullOrWhiteSpace(code.Number))
                {
                    var path = Path.Combine(currentDir, code.Number + PdfExtension);
                    if (File.Exists(path))
                        return true;
                }

                // Check latest code path
                if (!string.IsNullOrWhiteSpace(code.LatestCode))
                {
                    var path = Path.Combine(currentDir, code.LatestCode + PdfExtension);
                    if (File.Exists(path))
                        return true;
                }
            }
            catch { /* Best effort attempt */ }

            return false;
        }

        private static async Task<bool> TryHeadIsPdfAsync(HttpClient http, Uri uri, CancellationToken token)
        {
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Head, uri);
                var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);

                if (!resp.IsSuccessStatusCode)
                    return false;

                var media = resp.Content.Headers.ContentType?.MediaType;
                return string.Equals(media, "application/pdf", StringComparison.OrdinalIgnoreCase) ||
                       uri.AbsolutePath.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase);
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Helper Methods - PDF Processing

        private async Task ExtractCodeTextToQueueAsync(
            CodeItem code,
            HttpClient http,
            SemaphoreSlim extractionSemaphore,
            ChannelWriter<(CodeItem code, string text)> writer,
            CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return;

            await extractionSemaphore.WaitAsync(token).ConfigureAwait(false);
            try
            {
                // If the row has no link, resolve the PDF URL from the configured settings links.
                var link = code.Link;
                if (string.IsNullOrWhiteSpace(link))
                {
                    var settingsList = Settings?.ToList() ?? new List<SettingEntry>();
                    var mapping = await TryFindMappedPdfAsync(http, code.Number ?? string.Empty, settingsList, token).ConfigureAwait(false);
                    if (mapping.found && !string.IsNullOrWhiteSpace(mapping.foundLink))
                    {
                        link = mapping.foundLink;
                        SafeDispatch(() => code.Link = link);
                    }
                    else
                    {
                        AppendConsole($"Skipping {code.Number}: no link.");
                        UpdateCodeWithoutPages(code);
                        AdvanceProgress();
                        return;
                    }
                }

                var pages = await TryExtractFirstPageTextAsync(link, http, token).ConfigureAwait(false);
                if (pages == null || pages.Length == 0)
                {
                    UpdateCodeWithoutPages(code);
                    AppendConsole($"No text extracted for {code.Number} (first page empty or not available).");
                    AdvanceProgress();
                    return;
                }

                await writer.WriteAsync((code, pages[0] ?? string.Empty), token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                MarkPdfCheckTimedOut(code);
                AppendConsole($"Skipped {code.Number}: PDF URL timed out.");
                AdvanceProgress();
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                AppendConsole($"Error extracting {code.Number}: {ex.Message}");
            }
            finally
            {
                extractionSemaphore.Release();
            }
        }

        private static async Task<bool> TryUrlHasPdfAsync(HttpClient http, Uri uri, CancellationToken token)
        {
            try
            {
                await PdfTextExtractor.DownloadPdfBytesFromUrlAsync(http, uri.AbsoluteUri, token).ConfigureAwait(false);
                return true;
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                throw;
            }
            catch (HttpRequestException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private async Task ConsumeExtractedTextAsync(
            ChannelReader<(CodeItem code, string text)> reader,
            CancellationToken token)
        {
            await foreach (var item in reader.ReadAllAsync(token).ConfigureAwait(false))
            {
                if (token.IsCancellationRequested)
                    break;

                try
                {
                    var parsed = await ParsePdfTextAsync(item.code, item.text, token).ConfigureAwait(false);
                    ApplyParsedPdfInfo(item.code, parsed);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    AppendConsole($"Error processing {item.code.Number}: {ex.Message}");
                }
                finally
                {
                    AdvanceProgress();
                }
            }
        }

        private async Task ProcessCodeUpdateAsync(CodeItem code, HttpClient http, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(code.Link))
            {
                AppendConsole($"Skipping {code.Number}: no link.");
                return;
            }

            try
            {
                var pages = await TryExtractFirstPageTextAsync(code.Link, http, token).ConfigureAwait(false);
                if (pages == null || pages.Length == 0)
                {
                    UpdateCodeWithoutPages(code);
                    AppendConsole($"No text extracted for {code.Number} (first page empty or not available).");
                    return;
                }

                var parsed = await ParsePdfTextAsync(code, pages[0] ?? string.Empty, token).ConfigureAwait(false);
                ApplyParsedPdfInfo(code, parsed);
            }
            catch (OperationCanceledException)
            {
                AppendConsole("Search cancelled during download.");
                throw;
            }
            catch (Exception ex)
            {
                AppendConsole($"Error processing {code.Number}: {ex.Message}");
            }
        }

        private async Task ProcessLocalCodeUpdateAsync(
            CodeItem code,
            string pdfFolderForTab,
            SemaphoreSlim semaphore,
            CancellationToken token)
        {
            if (token.IsCancellationRequested)
                return;

            await semaphore.WaitAsync(token).ConfigureAwait(false);
            try
            {
                if (string.IsNullOrWhiteSpace(code.Number))
                {
                    AppendConsole("Skipping row: no code number configured.");
                    return;
                }

                var foundPath = FindLocalPdfPath(code, pdfFolderForTab);
                if (foundPath == null)
                {
                    UpdateCodeWithoutPages(code);
                    return;
                }

                AppendConsole($"Reading local PDF: {Path.GetFileName(foundPath)} for {code.Number}");

                var pages = await Task.Run(() => PdfTextExtractor.ExtractTextPerPage(foundPath), token).ConfigureAwait(false);
                if (pages == null || pages.Length == 0)
                {
                    UpdateCodeWithoutPages(code);
                    AppendConsole($"No text extracted from {Path.GetFileName(foundPath)} for {code.Number}.");
                    return;
                }

                var parsed = await ParsePdfTextAsync(code, pages[0] ?? string.Empty, token).ConfigureAwait(false);
                ApplyParsedPdfInfo(code, parsed);
                AppendConsole($"Successfully updated {code.Number} from local PDF");
            }
            catch (OperationCanceledException)
            {
                AppendConsole($"Processing cancelled for {code.Number}");
            }
            catch (Exception ex)
            {
                AppendConsole($"Error processing {code.Number}: {ex.Message}");
            }
            finally
            {
                semaphore.Release();
            }
        }

        private string? FindLocalPdfPath(CodeItem code, string pdfFolderForTab)
        {
            var candidatePaths = new List<string>();

            // Add number-based path
            var safeNumber = MakeSafeFileName(code.Number);
            if (!safeNumber.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                safeNumber += PdfExtension;
            candidatePaths.Add(Path.Combine(pdfFolderForTab, safeNumber));

            // Add latest code path if different
            if (!string.IsNullOrWhiteSpace(code.LatestCode) &&
                !string.Equals(code.Number, code.LatestCode, StringComparison.OrdinalIgnoreCase))
            {
                var safeLatest = MakeSafeFileName(code.LatestCode);
                if (!safeLatest.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                    safeLatest += PdfExtension;
                candidatePaths.Add(Path.Combine(pdfFolderForTab, safeLatest));
            }

            foreach (var candidate in candidatePaths)
            {
                if (File.Exists(candidate))
                    return candidate;
            }

            AppendConsole($"No local PDF found for {code.Number} (tried: {string.Join(", ", candidatePaths.Select(Path.GetFileName))})");
            return null;
        }

        private void UpdateCodeWithoutPages(CodeItem code)
        {
            SafeDispatch(() =>
            {
                code.HasCheck = true;
                code.HasUpdate = false;
                code.CodeExists = false;
            });
        }

        private async Task<PdfTextComparer.PdfCodeInfo?> ParsePdfTextAsync(CodeItem code, string firstPageText, CancellationToken token)
        {
            var matchedSetting = GetMatchingSettingForCode(code.Number);
            var isIcc = IsIccSource(code, matchedSetting);
            var contextSettings = ContextSettings.FirstOrDefault() ?? new ExtractionContextSettings();
            var latestContexts = SplitContextText(contextSettings.LatestCode);
            var issueContexts = SplitContextText(contextSettings.IssueDate);
            var expirationContexts = SplitContextText(contextSettings.ExpirationDate);
            var parsed = isIcc
                ? PdfTextComparer.ParseIccEs(firstPageText, latestContexts, issueContexts, expirationContexts)
                : PdfTextComparer.ParseIapmo(firstPageText, latestContexts, issueContexts, expirationContexts);

            // Run the (slow) AI model according to the selected mode:
            //  - Ai:     AI is authoritative and checks every field; Regex is only a fallback.
            //  - Hybrid: Regex is authoritative; AI only fills fields Regex could not find.
            //  - NonAi:  no AI at all.
            var useAi = ExtractionMode != AiMode.NonAi;
            var regexMissing = IsMissingValue(parsed.LatestCode) ||
                               IsMissingValue(parsed.IssueDate) ||
                               IsMissingValue(parsed.ExpirationDate);
            var codeKey = code.Number?.Trim() ?? string.Empty;
            var productsResolvableWithoutAi = !CalculateProduct ||
                                              (codeKey.Length > 0 && _productCountCache.ContainsKey(codeKey)) ||
                                              CountProductsFromDescription(code.Description).HasValue;

            // In Ai mode we always consult the model; in Hybrid mode only when it can help.
            var needsAi = ExtractionMode == AiMode.Ai || regexMissing || !productsResolvableWithoutAi;

            if (useAi && needsAi)
            {
                try
                {
                    var modelFile = string.IsNullOrWhiteSpace(AiModelFileName) ? DefaultAiModelFileName : AiModelFileName;
                    if (_aiExtractor == null || _aiModelFileNameUsed != modelFile)
                    {
                        _aiExtractor?.Dispose();
                        _aiExtractor = new AiPdfTextExtractor(Path.Combine(AppContext.BaseDirectory, "AIModels", modelFile), CpuPercent);
                        _aiModelFileNameUsed = modelFile;
                    }
                    var aiContexts = latestContexts.Concat(issueContexts).Concat(expirationContexts).ToArray();
                    var ai = await _aiExtractor.ExtractAsync(firstPageText, aiContexts, token, CalculateProduct).ConfigureAwait(false);
                    if (ai != null)
                    {
                        if (ExtractionMode == AiMode.Ai)
                        {
                            // AI is authoritative; Regex only fills what AI could not determine.
                            if (!IsMissingValue(ai.LatestCode))
                                parsed.LatestCode = ai.LatestCode;
                            if (!IsMissingValue(ai.IssueDate))
                                parsed.IssueDate = ai.IssueDate;
                            if (!IsMissingValue(ai.ExpirationDate))
                                parsed.ExpirationDate = ai.ExpirationDate;
                        }
                        else // Hybrid - Regex is authoritative; AI only fills missing fields.
                        {
                            if (IsMissingValue(parsed.LatestCode) && !IsMissingValue(ai.LatestCode))
                                parsed.LatestCode = ai.LatestCode;
                            if (IsMissingValue(parsed.IssueDate) && !IsMissingValue(ai.IssueDate))
                                parsed.IssueDate = ai.IssueDate;
                            if (IsMissingValue(parsed.ExpirationDate) && !IsMissingValue(ai.ExpirationDate))
                                parsed.ExpirationDate = ai.ExpirationDate;
                        }

                        if (CalculateProduct && ai.ProductsCount.HasValue)
                            parsed.ProductsCount = ai.ProductsCount.Value;
                    }
                }
                catch (Exception ex)
                {
                    AppendConsole($"AI extraction unavailable; Regex result kept: {ex.Message}");
                }
            }

            if (CalculateProduct)
                parsed.ProductsCount = ResolveProductCount(code, parsed);

            return parsed;
        }

        // Counts products for a report: reuse the learned cache value, otherwise count the
        // product codes from the Description, and only fall back to the AI result. The outcome
        // is remembered so the next run reuses it instead of recomputing.
        private int? ResolveProductCount(CodeItem code, PdfTextComparer.PdfCodeInfo parsed)
        {
            var key = code.Number?.Trim();
            if (string.IsNullOrWhiteSpace(key))
                return parsed.ProductsCount;

            if (_productCountCache.TryGetValue(key, out var cached))
                return cached;

            var count = CountProductsFromDescription(code.Description);
            if (count == null || count <= 0)
                count = parsed.ProductsCount;

            if (count.HasValue)
            {
                _productCountCache[key] = count.Value;
                SaveProductCountCache();
            }

            return count;
        }

        private static int? CountProductsFromDescription(string? description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return null;

            var text = description;
            var colon = text.IndexOf(':');
            if (colon >= 0 && colon < text.Length)
                text = text.Substring(colon + 1);

            var items = text
                .Split(new[] { ',', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => x.Length > 0)
                .ToList();

            return items.Count > 0 ? items.Count : (int?)null;
        }

        private void ApplyParsedPdfInfo(CodeItem code, PdfTextComparer.PdfCodeInfo? parsed)
        {
            try
            {
                SafeDispatch(() =>
                {
                    code.LatestCode_Old = code.LatestCode;
                    code.IssueDate_Old = code.IssueDate;
                    code.ExpirationDate_Old = code.ExpirationDate;
                    code.ProductsListed_Old = code.ProductsListed;

                    code.LatestCode = parsed?.LatestCode ?? string.Empty;
                    code.IssueDate = parsed?.IssueDate ?? "n/a";
                    code.ExpirationDate = parsed?.ExpirationDate ?? "n/a";

                    if (CalculateProduct && parsed?.ProductsCount.HasValue == true)
                        code.ProductsListed = parsed.ProductsCount.Value.ToString();

                    code.HasCheck = true;
                    code.CodeExists = true;
                    code.IsPdfCheckTimedOut = false;
                    code.HasUpdate = HasValueChanged(code.LatestCode, code.LatestCode_Old) ||
                                     HasValueChanged(code.IssueDate, code.IssueDate_Old) ||
                                     HasValueChanged(code.ExpirationDate, code.ExpirationDate_Old) ||
                                     HasValueChanged(code.ProductsListed, code.ProductsListed_Old);
                });

                LogCodeChanges(code);

                if (CalculateProduct && parsed?.ProductsCount.HasValue == true)
                    AppendConsole($"{code.Number} products listed -> {parsed.ProductsCount.Value}");
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to update model for {code.Number}: {ex.Message}");
            }
        }

        private void LogCodeChanges(CodeItem code)
        {
            var changes = new List<string>();

            if (HasValueChanged(code.LatestCode, code.LatestCode_Old))
                changes.Add($"LatestCode: '{code.LatestCode_Old}' -> '{code.LatestCode}'");

            if (HasValueChanged(code.IssueDate, code.IssueDate_Old))
                changes.Add($"IssueDate: '{code.IssueDate_Old}' -> '{code.IssueDate}'");

            if (HasValueChanged(code.ExpirationDate, code.ExpirationDate_Old))
                changes.Add($"ExpirationDate: '{code.ExpirationDate_Old}' -> '{code.ExpirationDate}'");

            if (changes.Count == 0)
                AppendConsole($"No changes detected for {code.Number} (first page parsed).");
            else
                changes.ForEach(c => AppendConsole($"{code.Number} updated: {c}"));
        }

        private static bool IsMissingValue(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ||
                   string.Equals(value.Trim(), "n/a", StringComparison.OrdinalIgnoreCase);
        }

        private static string[] SplitContextText(string? value)
        {
            return (value ?? string.Empty)
                .Split(new[] { ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(context => context.Trim())
                .Where(context => context.Length > 0)
                .ToArray();
        }

        private void EnsureContextDefaults()
        {
            var contexts = ContextSettings.FirstOrDefault();
            if (contexts == null)
            {
                ContextSettings.Add(new ExtractionContextSettings());
                contexts = ContextSettings[0];
            }

            if (string.IsNullOrWhiteSpace(contexts.LatestCode))
                contexts.LatestCode = DefaultLatestCodeContexts;
            if (string.IsNullOrWhiteSpace(contexts.IssueDate))
                contexts.IssueDate = DefaultIssueDateContexts;
            if (string.IsNullOrWhiteSpace(contexts.ExpirationDate))
                contexts.ExpirationDate = DefaultExpirationDateContexts;
        }

        private static bool HasValueChanged(string? current, string? previous)
        {
            var normalizedCurrent = IsMissingValue(current) ? "n/a" : current!.Trim();
            var normalizedPrevious = IsMissingValue(previous) ? "n/a" : previous!.Trim();
            return !string.Equals(normalizedCurrent, normalizedPrevious, StringComparison.OrdinalIgnoreCase);
        }

        private async Task<string[]?> TryExtractFirstPageTextAsync(string link, HttpClient? http, CancellationToken token)
        {
            try
            {
                if (Uri.IsWellFormedUriString(link, UriKind.Absolute))
                {
                    var uri = new Uri(link);
                    if (uri.IsFile)
                    {
                        return File.Exists(uri.LocalPath) ? PdfTextExtractor.ExtractTextPerPage(uri.LocalPath) : null;
                    }

                    if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                    {
                        if (http != null)
                            return await PdfTextExtractor.ExtractTextPerPageFromUrlAsync(link, http, token).ConfigureAwait(false);

                        return await PdfTextExtractor.ExtractTextPerPageFromUrlAsync(link, token).ConfigureAwait(false);
                    }
                }
                else if (File.Exists(link))
                {
                    return PdfTextExtractor.ExtractTextPerPage(link);
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch { /* Best effort attempt */ }

            return null;
        }

        #endregion

        #region Helper Methods - Download

        private string? PrepareDestinationFolder(TabViewModel selTab, string? destBase)
        {
            var tabName = string.IsNullOrWhiteSpace(selTab.Header) ? "Unknown" : selTab.Header;
            var safeTabName = MakeSafeFileName(tabName);

            var baseDir = destBase ?? (AppContext.BaseDirectory ?? Directory.GetCurrentDirectory());
            var destFolder = Path.Combine(baseDir, PdfFolderName, safeTabName);

            try
            {
                Directory.CreateDirectory(destFolder);
                return destFolder;
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to create destination folder '{destFolder}': {ex.Message}");
                return null;
            }
        }

        private void ResetDownloadProgress(List<CodeItem> rowList)
        {
            try
            {
                SafeDispatch(() =>
                {
                    foreach (var r in rowList)
                    {
                        try { r.DownloadProcess = 0; } catch { /* Best effort */ }
                    }
                });
            }
            catch { /* Best effort */ }
        }

        private async Task<(CodeItem code, bool success, string message)> DownloadSinglePdfAsync(
            CodeItem code,
            string destFolder,
            HttpClient http,
            SemaphoreSlim semaphore,
            CancellationToken token)
        {
            if (token.IsCancellationRequested)
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, "Canceled");
            }

            await semaphore.WaitAsync(token).ConfigureAwait(false);
            try
            {
                if (string.IsNullOrWhiteSpace(code.Link))
                {
                    SafeDispatch(() => code.DownloadProcess = 0);
                    return (code, false, "No link");
                }

                if (Uri.IsWellFormedUriString(code.Link, UriKind.Absolute))
                {
                    var uri = new Uri(code.Link);

                    if (uri.IsFile)
                        return await DownloadFromFileAsync(code, uri, destFolder).ConfigureAwait(false);

                    if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                        return await DownloadFromHttpAsync(code, uri, destFolder, http, token).ConfigureAwait(false);
                }

                // Fallback to local file path
                return await DownloadFromLocalPathAsync(code, destFolder).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, "Canceled");
            }
            catch (Exception ex)
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, ex.Message);
            }
            finally
            {
                semaphore.Release();
            }
        }

        private async Task<(CodeItem code, bool success, string message)> DownloadFromFileAsync(
            CodeItem code,
            Uri uri,
            string destFolder)
        {
            var local = uri.LocalPath;
            if (!File.Exists(local))
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, $"Source file not found: {local}");
            }

            var fileName = MakeSafeFileName(code.Number ?? Path.GetFileName(local));
            if (!fileName.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                fileName += PdfExtension;

            var target = Path.Combine(destFolder, fileName);
            await Task.Run(() => File.Copy(local, target, overwrite: true)).ConfigureAwait(false);

            SafeDispatch(() => code.DownloadProcess = 100);
            return (code, true, target);
        }

        private async Task<(CodeItem code, bool success, string message)> DownloadFromHttpAsync(
            CodeItem code,
            Uri uri,
            string destFolder,
            HttpClient http,
            CancellationToken token)
        {
            var fileName = GenerateFileName(code, uri);
            var target = Path.Combine(destFolder, fileName);

            // The source may be an HTML viewer page rather than the PDF endpoint.
            var bytes = await PdfTextExtractor.DownloadPdfBytesFromUrlAsync(uri.AbsoluteUri, token).ConfigureAwait(false);
            await File.WriteAllBytesAsync(target, bytes, token).ConfigureAwait(false);

            SafeDispatch(() => code.DownloadProcess = 100);
            return (code, true, target);
        }

        private async Task<(CodeItem code, bool success, string message)> DownloadFromLocalPathAsync(
            CodeItem code,
            string destFolder)
        {
            var localPath = code.Link!;
            if (!File.Exists(localPath))
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, $"Source file not found: {localPath}");
            }

            var fileName = MakeSafeFileName(code.Number ?? Path.GetFileName(localPath));
            if (!fileName.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                fileName += PdfExtension;

            var target = Path.Combine(destFolder, fileName);
            await Task.Run(() => File.Copy(localPath, target, overwrite: true)).ConfigureAwait(false);

            SafeDispatch(() => code.DownloadProcess = 100);
            return (code, true, target);
        }

        private static string GenerateFileName(CodeItem code, Uri uri)
        {
            var suggested = code.Number;
            if (string.IsNullOrWhiteSpace(suggested))
                suggested = Path.GetFileName(uri.LocalPath);
            if (string.IsNullOrWhiteSpace(suggested))
                suggested = Guid.NewGuid().ToString();

            var fileName = MakeSafeFileName(suggested);
            if (!fileName.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                fileName += PdfExtension;

            return fileName;
        }

        private static bool IsPdfContent(byte[] buffer, int bytesRead)
        {
            if (bytesRead < 4)
                return false;

            var signature = System.Text.Encoding.ASCII.GetBytes("%PDF");
            for (int i = 0; i <= bytesRead - signature.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < signature.Length; j++)
                {
                    if (buffer[i + j] != signature[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return true;
            }

            return false;
        }

        private async Task WriteStreamToFileAsync(
            Stream responseStream,
            string target,
            byte[] sniffBuffer,
            int sniffRead,
            long? totalBytes,
            CodeItem code,
            CancellationToken token)
        {
            var buffer = new byte[FileBufferSize];
            long copied = 0;

            using var fs = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None, FileBufferSize, useAsync: true);

            // Write sniffed data first
            if (sniffRead > 0)
            {
                await fs.WriteAsync(sniffBuffer, 0, sniffRead, token).ConfigureAwait(false);
                copied += sniffRead;
            }

            // Write remaining data with progress
            int read;
            while ((read = await responseStream.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)) > 0)
            {
                await fs.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
                copied += read;

                if (totalBytes.HasValue && totalBytes.Value > 0)
                {
                    double progress = (double)copied / totalBytes.Value;
                    SafeDispatch(() => code.DownloadProcess = (int)Math.Round(progress * 100));
                }
            }
        }

        private void ProcessDownloadResults(IEnumerable<(CodeItem code, bool success, string message)> results)
        {
            foreach (var (code, success, message) in results)
            {
                if (success)
                {
                    SafeDispatch(() =>
                    {
                        code.CodeExists = true;
                        code.HasUpdate = true;
                        code.LastCheck = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
                    });

                    AppendConsole($"Downloaded {code.Number ?? "(no number)"} -> {message}");
                }
                else
                {
                    AppendConsole($"Failed to download {code.Number ?? "(no number)"}: {message}");
                }

                SafeDispatch(() => code.DownloadProcess = success ? 100 : 0);
            }
        }

        #endregion

        #region Helper Methods - Settings & Mapping

        private static async Task<(bool found, string? foundLink, string? settingName)> TryFindMappedPdfAsync(
            HttpClient http,
            string number,
            List<SettingEntry> settings,
            CancellationToken token)
        {
            var (prefix, numeric) = ParseCodeNumber(number.Trim());

            var matched = FindMatchingSetting(settings, number.Trim(), prefix);
            if (matched == null)
                return (false, null, null);

            var stems = BuildSearchStems(number.Trim(), numeric);
            var basesToTry = new List<string>();

            if (!string.IsNullOrWhiteSpace(matched.Link))
                basesToTry.Add(matched.Link);
            if (!string.IsNullOrWhiteSpace(matched.PdfFolder))
                basesToTry.Add(matched.PdfFolder);

            foreach (var baseStr in basesToTry)
            {
                if (token.IsCancellationRequested)
                    break;

                if (string.IsNullOrWhiteSpace(baseStr))
                    continue;

                foreach (var stem in stems)
                {
                    if (token.IsCancellationRequested)
                        break;

                    var fileName = stem.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase) ? stem : stem + PdfExtension;
                    var candidateUri = CombineUriBase(baseStr, fileName);

                    if (candidateUri == null)
                        continue;

                    try
                    {
                        using var req = new HttpRequestMessage(HttpMethod.Head, candidateUri);
                        var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);

                        if (resp.IsSuccessStatusCode)
                        {
                            var media = resp.Content.Headers.ContentType?.MediaType;
                            if (string.Equals(media, "application/pdf", StringComparison.OrdinalIgnoreCase) ||
                                candidateUri.AbsolutePath.EndsWith(PdfExtension, StringComparison.OrdinalIgnoreCase))
                            {
                                return (true, candidateUri.ToString(), matched.Name);
                            }
                        }
                    }
                    catch { /* Try next candidate */ }
                }
            }

            return (false, null, null);
        }

        private static (string prefix, string numeric) ParseCodeNumber(string trimmed)
        {
            var match = Regex.Match(trimmed, @"^(?<pre>[A-Za-z]+)[\s\-]*0*(?<num>\d+)$", RegexOptions.Compiled);
            if (match.Success)
                return (match.Groups["pre"].Value, match.Groups["num"].Value);

            var parts = trimmed.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2 && Regex.IsMatch(parts[0], @"^[A-Za-z]+$"))
            {
                var prefix = parts[0];
                var numeric = string.Concat(parts.Skip(1)).TrimStart('0');
                return (prefix, string.IsNullOrWhiteSpace(numeric) ? parts.Last() : numeric);
            }

            return (string.Empty, string.Empty);
        }

        private static SettingEntry? FindMatchingSetting(List<SettingEntry> settings, string trimmed, string prefix)
        {
            if (!string.IsNullOrWhiteSpace(prefix))
            {
                var exact = settings.FirstOrDefault(s =>
                    string.Equals(s.Type?.Trim(), prefix.Trim(), StringComparison.OrdinalIgnoreCase));
                if (exact != null)
                    return exact;
            }

            return settings.FirstOrDefault(s =>
                !string.IsNullOrWhiteSpace(s.Type) &&
                trimmed.Contains(s.Type, StringComparison.OrdinalIgnoreCase));
        }

        private static List<string> BuildSearchStems(string trimmed, string numeric)
        {
            var stems = new List<string>();

            var cleaned = Regex.Replace(trimmed, @"[\s\-]+", "", RegexOptions.Compiled);
            if (!string.IsNullOrWhiteSpace(cleaned))
                stems.Add(cleaned);

            if (!string.IsNullOrWhiteSpace(numeric))
            {
                stems.Add(numeric);
                if (numeric.Length == 3)
                    stems.Add("0" + numeric);
            }

            return stems.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private SettingEntry? GetMatchingSettingForCode(string? codeNumber)
        {
            if (string.IsNullOrWhiteSpace(codeNumber))
                return null;

            var (prefix, _) = ParseCodeNumber(codeNumber.Trim());

            if (!string.IsNullOrWhiteSpace(prefix))
            {
                var exact = Settings.FirstOrDefault(s =>
                    string.Equals(s.Type?.Trim(), prefix.Trim(), StringComparison.OrdinalIgnoreCase));
                if (exact != null)
                    return exact;
            }

            return Settings.FirstOrDefault(s =>
                !string.IsNullOrWhiteSpace(s.Type) &&
                codeNumber.Contains(s.Type, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsIccSource(CodeItem code, SettingEntry? matchedSetting)
        {
            if (!string.IsNullOrWhiteSpace(code.WebType) && code.WebType.Contains("icc", StringComparison.OrdinalIgnoreCase))
                return true;

            if (!string.IsNullOrWhiteSpace(code.Link) && code.Link.Contains("icc-es.org", StringComparison.OrdinalIgnoreCase))
                return true;

            if (matchedSetting != null &&
                ((matchedSetting.Type?.Contains("ESR", StringComparison.OrdinalIgnoreCase) ?? false) ||
                 (matchedSetting.Name?.Contains("ICC", StringComparison.OrdinalIgnoreCase) ?? false)))
                return true;

            return false;
        }

        #endregion

        #region Helper Methods - Utility

        private string GetPdfFolderForTab(TabViewModel selTab)
        {
            var tabName = string.IsNullOrWhiteSpace(selTab.Header) ? "Unknown" : selTab.Header;
            var safeTabName = MakeSafeFileName(tabName);
            var baseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            return Path.Combine(baseDir, PdfFolderName, safeTabName);
        }

        private static int CalculateMaxConcurrency(int multiplier)
        {
            var cpuWorkerLimit = Math.Max(1, (int)Math.Ceiling(Environment.ProcessorCount * CpuLimitPercent / 100d));
            var requested = Environment.ProcessorCount * multiplier;
            return Math.Clamp(requested, 1, cpuWorkerLimit);
        }

        private void ResetPdfTimeoutState(IEnumerable<CodeItem> rows)
        {
            SafeDispatch(() =>
            {
                foreach (var code in rows)
                    code.IsPdfCheckTimedOut = false;
            });
        }

        private void MarkPdfCheckTimedOut(CodeItem code)
        {
            SafeDispatch(() =>
            {
                code.IsPdfCheckTimedOut = true;
                code.CodeExists = false;
                code.HasCheck = true;
                code.HasUpdate = false;
                code.LastCheck = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
            });
        }

        private void SafeDispatch(Action action)
        {
            try
            {
                Application.Current?.Dispatcher.Invoke(action);
            }
            catch { /* Best effort attempt */ }
        }

        private static string MakeSafeFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return DefaultFileName;

            var invalid = Path.GetInvalidFileNameChars();
            var chars = input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
            var safe = new string(chars);

            safe = Regex.Replace(safe, @"\s{2,}", " ").Trim();
            safe = Regex.Replace(safe, @"[\. ]+$", "");

            return string.IsNullOrEmpty(safe) ? DefaultFileName : safe;
        }

        private static Uri? CombineUriBase(string baseUrl, string fileName)
        {
            if (string.IsNullOrWhiteSpace(baseUrl) || string.IsNullOrWhiteSpace(fileName))
                return null;

            try
            {
                var normalizedBase = baseUrl.Trim();
                if (!normalizedBase.EndsWith("/", StringComparison.Ordinal))
                    normalizedBase += "/";

                if (Uri.TryCreate(normalizedBase, UriKind.Absolute, out var baseUri))
                {
                    var escaped = Uri.EscapeDataString(fileName);
                    return new Uri(baseUri, escaped);
                }

                var prefixed = normalizedBase.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                    ? normalizedBase
                    : "http://" + normalizedBase;

                if (Uri.TryCreate(prefixed, UriKind.Absolute, out baseUri))
                {
                    var escaped = Uri.EscapeDataString(fileName);
                    return new Uri(baseUri, escaped);
                }
            }
            catch { /* Best effort attempt */ }

            return null;
        }

        #endregion

        #region Nested Classes

        public enum ConsoleLevel
        {
            Info,
            Success,
            Warning,
            Error
        }

        public enum AiMode
        {
            Ai,
            NonAi,
            Hybrid
        }

        private static AiMode ParseExtractionMode(string? value)
        {
            if (Enum.TryParse<AiMode>(value, true, out var mode))
                return mode;
            return AiMode.Ai;
        }

        public sealed class ConsoleLine
        {
            public string Text { get; }
            public ConsoleLevel Level { get; }

            public ConsoleLine(string text, ConsoleLevel level)
            {
                Text = text;
                Level = level;
            }
        }

        private class AppSettings
        {
            public List<SettingEntry>? WebSettings { get; set; }
            public ExtractionContextSettings? ContextSettings { get; set; }
            public string? ExtractionMode { get; set; }
            public bool? CalculateProduct { get; set; }
            public int? CpuPercent { get; set; }
            public string? AiModelFileName { get; set; }
        }

        #endregion
    }
}
