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
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace CodeReportTracker.ViewModels
{
    public partial class MainWindowViewModel : ObservableObject
    {
        #region Constants

        private const int DefaultHttpTimeoutSeconds = 10;
        private const int ExtendedHttpTimeoutSeconds = 60;
        private const int SniffBufferSize = 8192;
        private const int FileBufferSize = 81920;
        private const string DefaultFileName = "Unnamed";
        private const string PdfExtension = ".pdf";
        private const string PdfFolderName = "Pdf Files";
        private const string SettingsFileName = "settings.json";

        #endregion

        #region Fields

        private string _consoleText = string.Empty;
        private bool _isBusy;
        private TabViewModel? _selectedTab;
        private readonly string _currentDir = Directory.GetCurrentDirectory();

        // Delegates provided by the View for view-only operations
        private readonly Func<Task>? _searchAction;
        private readonly Action? _stopAction;
        private readonly Action? _selectExcelAction;
        private readonly Func<Task>? _exportAction;

        #endregion

        #region Properties

        public ObservableCollection<SettingEntry> Settings { get; }
        public ObservableCollection<TabViewModel> Tabs { get; }

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

        #endregion

        #region Constructor

        public MainWindowViewModel(
            Func<Task>? searchAction = null,
            Action? stopAction = null,
            Action? selectExcelAction = null,
            Func<Task>? exportAction = null)
        {
            _searchAction = searchAction;
            _stopAction = stopAction;
            _selectExcelAction = selectExcelAction;
            _exportAction = exportAction;

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
        }

        #endregion

        #region Public Methods

        public void AppendConsole(string line)
        {
            ConsoleText = string.IsNullOrEmpty(ConsoleText)
                ? line + Environment.NewLine
                : ConsoleText + line + Environment.NewLine;
        }

        public void SaveSettings()
        {
            try
            {
                var settingsPath = GetSettingsPath();
                var appSettings = new AppSettings
                {
                    WebSettings = new List<SettingEntry>(Settings)
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

                AppendConsole($"Settings loaded from: {settingsPath}");
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to load settings: {ex.Message}");
                AppendConsole("Using default settings.");
            }
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

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(ExtendedHttpTimeoutSeconds) };

            foreach (var code in rowList)
            {
                if (token.IsCancellationRequested)
                {
                    AppendConsole("Search cancelled.");
                    break;
                }

                await ProcessCodeUpdateAsync(code, http, token).ConfigureAwait(false);
            }

            AppendConsole("Search finished.");
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
            var results = await Task.WhenAll(tasks).ConfigureAwait(false);

            ProcessDownloadResults(results);
            AppendConsole("Download finished.");
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

            await semaphore.WaitAsync(token).ConfigureAwait(false);
            try
            {
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
            catch (Exception ex)
            {
                AppendConsole($"Error checking {code.Number}: {ex.Message}");
            }
            finally
            {
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
                    await TryHeadIsPdfAsync(http, uri, token).ConfigureAwait(false))
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
            catch
            {
                return false;
            }
        }

        #endregion

        #region Helper Methods - PDF Processing

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

                var parsed = ParsePdfText(code, pages[0] ?? string.Empty);
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

                var parsed = await Task.Run(() => ParsePdfText(code, pages[0] ?? string.Empty), token).ConfigureAwait(false);
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
                code.HasUpdate = true;
                code.CodeExists = false;
                code.LastCheck = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
            });
        }

        private PdfTextComparer.PdfCodeInfo? ParsePdfText(CodeItem code, string firstPageText)
        {
            var matchedSetting = GetMatchingSettingForCode(code.Number);
            var isIcc = IsIccSource(code, matchedSetting);
            return isIcc ? PdfTextComparer.ParseIccEs(firstPageText) : PdfTextComparer.ParseIapmo(firstPageText);
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

                    code.LatestCode = parsed?.LatestCode ?? string.Empty;
                    code.IssueDate = parsed?.IssueDate ?? "n/a";
                    code.ExpirationDate = parsed?.ExpirationDate ?? "n/a";

                    code.CodeExists = true;
                    code.HasUpdate = true;
                    code.LastCheck = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
                });

                LogCodeChanges(code);
            }
            catch (Exception ex)
            {
                AppendConsole($"Failed to update model for {code.Number}: {ex.Message}");
            }
        }

        private void LogCodeChanges(CodeItem code)
        {
            var changes = new List<string>();

            if (!string.Equals(code.LatestCode, code.LatestCode_Old, StringComparison.OrdinalIgnoreCase))
                changes.Add($"LatestCode: '{code.LatestCode_Old}' -> '{code.LatestCode}'");

            if (!string.Equals(code.IssueDate, code.IssueDate_Old, StringComparison.OrdinalIgnoreCase))
                changes.Add($"IssueDate: '{code.IssueDate_Old}' -> '{code.IssueDate}'");

            if (!string.Equals(code.ExpirationDate, code.ExpirationDate_Old, StringComparison.OrdinalIgnoreCase))
                changes.Add($"ExpirationDate: '{code.ExpirationDate_Old}' -> '{code.ExpirationDate}'");

            if (changes.Count == 0)
                AppendConsole($"No changes detected for {code.Number} (first page parsed).");
            else
                changes.ForEach(c => AppendConsole($"{code.Number} updated: {c}"));
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
                        return await PdfTextExtractor.ExtractTextPerPageFromUrlAsync(link, token).ConfigureAwait(false);
                    }
                }
                else if (File.Exists(link))
                {
                    return PdfTextExtractor.ExtractTextPerPage(link);
                }
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
            using var req = new HttpRequestMessage(HttpMethod.Get, uri);
            using var resp = await http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);

            if (!resp.IsSuccessStatusCode)
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                return (code, false, $"HTTP {(int)resp.StatusCode}");
            }

            var fileName = GenerateFileName(code, uri);
            var target = Path.Combine(destFolder, fileName);

            using var responseStream = await resp.Content.ReadAsStreamAsync(token).ConfigureAwait(false);

            // Sniff content to verify PDF
            var sniffBuffer = new byte[SniffBufferSize];
            int sniffRead = await responseStream.ReadAsync(sniffBuffer, 0, sniffBuffer.Length, token).ConfigureAwait(false);

            if (!IsPdfContent(sniffBuffer, sniffRead))
            {
                SafeDispatch(() => code.DownloadProcess = 0);
                var ct = resp.Content.Headers.ContentType?.MediaType ?? "(no content-type)";
                return (code, false, $"Downloaded content is not a PDF. Content-Type: {ct}");
            }

            // Write to file with progress tracking
            await WriteStreamToFileAsync(responseStream, target, sniffBuffer, sniffRead, resp.Content.Headers.ContentLength, code, token).ConfigureAwait(false);

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
            return Math.Clamp(Environment.ProcessorCount * multiplier, 4, multiplier == 4 ? 32 : 16);
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
                    var escaped = Uri.EscapeUriString(fileName);
                    return new Uri(baseUri, escaped);
                }

                var prefixed = normalizedBase.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                    ? normalizedBase
                    : "http://" + normalizedBase;

                if (Uri.TryCreate(prefixed, UriKind.Absolute, out baseUri))
                {
                    var escaped = Uri.EscapeUriString(fileName);
                    return new Uri(baseUri, escaped);
                }
            }
            catch { /* Best effort attempt */ }

            return null;
        }

        #endregion

        #region Nested Classes

        private class AppSettings
        {
            public List<SettingEntry>? WebSettings { get; set; }
        }

        #endregion
    }
}