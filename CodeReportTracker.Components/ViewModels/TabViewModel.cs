using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;
using CodeReportTracker.Core.Models;

namespace CodeReportTracker.Components.ViewModels
{
    public class TabViewModel : INotifyPropertyChanged
    {
        private string _header = "New Tab";
        private object? _content;
        private bool _isEditing;
        private object? _selectedItem;
        private string _pdfFolder = string.Empty;

        public TabViewModel() { }

        public TabViewModel(string header, object? content = null)
        {
            Header = header;
            Content = content;
        }

        public string Header
        {
            get => _header;
            set
            {
                if (_header == value) return;
                _header = value;
                OnPropertyChanged();
            }
        }

        // Optional: can hold arbitrary content (view-model or other data)
        public object? Content
        {
            get => _content;
            set
            {
                if (_content == value) return;
                _content = value;
                OnPropertyChanged();
            }
        }

        // Per-tab data collection. Use CodeItem so CodeReportTable can bind to it directly
        // and avoid wrapping the source into a new collection.
        public ObservableCollection<CodeItem> Items { get; } = new ObservableCollection<CodeItem>();

        // Per-tab selected item. Bind CodeReportTable.SelectedItem to this so selection is preserved per-tab.
        public object? SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem == value) return;
                _selectedItem = value;
                OnPropertyChanged();
            }
        }

        // Used by header template to toggle inline edit mode
        public bool IsEditing
        {
            get => _isEditing;
            set
            {
                if (_isEditing == value) return;
                _isEditing = value;
                OnPropertyChanged();
            }
        }

        // Optional per-tab commands (host can also handle at collection level)
        public ICommand? CloseCommand { get; set; }
        public ICommand? RenameCommand { get; set; }

        // Per-tab PDF folder path (computed from Header via InitializePdfFolder)
        public string PdfFolder
        {
            get => _pdfFolder;
            private set
            {
                if (_pdfFolder == value) return;
                _pdfFolder = value;
                OnPropertyChanged();
            }
        }

        // Cached list of PDF file paths found in PdfFolder
        public ObservableCollection<string> PdfFiles { get; } = new ObservableCollection<string>();

        /// <summary>
        /// Initialize the PdfFolder for this tab using the provided base directory (or AppContext/BaseDirectory fallback).
        /// Creates the folder if it does not exist.
        /// </summary>
        /// <param name="baseDir">Optional base directory. If null uses AppContext.BaseDirectory or Directory.GetCurrentDirectory().</param>
        public void InitializePdfFolder(string? baseDir = null)
        {
            var root = baseDir ?? AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();
            var safeTabName = MakeSafeFileName(string.IsNullOrWhiteSpace(Header) ? "Unknown" : Header);
            var folder = Path.Combine(root, "Pdf Files", safeTabName);
            try
            {
                Directory.CreateDirectory(folder);
            }
            catch
            {
                // swallow - caller can check PdfFolder existence if required
            }

            PdfFolder = folder;
        }

        /// <summary>
        /// Ensures the PdfFolder has been initialized and exists. If not initialized it will call InitializePdfFolder.
        /// </summary>
        /// <param name="baseDir">Optional base directory used when initializing.</param>
        /// <returns>True when folder exists or was created; false otherwise.</returns>
        public bool EnsurePdfFolderExists(string? baseDir = null)
        {
            if (string.IsNullOrWhiteSpace(PdfFolder))
                InitializePdfFolder(baseDir);

            try
            {
                if (!Directory.Exists(PdfFolder))
                    Directory.CreateDirectory(PdfFolder);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Loads the list of PDF file paths from the folder into the PdfFiles collection.
        /// Caller may call EnsurePdfFolderExists first; this method will attempt to initialize if needed.
        /// </summary>
        /// <param name="baseDir">Optional base directory used when initializing the folder.</param>
        public Task LoadPdfFilesAsync(string? baseDir = null)
        {
            if (string.IsNullOrWhiteSpace(PdfFolder))
                InitializePdfFolder(baseDir);

            PdfFiles.Clear();

            try
            {
                if (!Directory.Exists(PdfFolder))
                    return Task.CompletedTask;

                var files = Directory.GetFiles(PdfFolder, "*.pdf");
                foreach (var f in files.OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
                {
                    PdfFiles.Add(f);
                }
            }
            catch
            {
                // ignore errors while loading file list
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Attempts to resolve a candidate local PDF file path for a given CodeItem using the same rules as download/local-update logic:
        /// prefer CodeItem.Number, then CodeItem.LatestCode, then filename extracted from CodeItem.Link. Returns full path or null.
        /// </summary>
        /// <param name="code">CodeItem to resolve</param>
        /// <returns>Full path to a matching PDF in PdfFolder, or null if not found.</returns>
        public string? GetCandidatePdfPath(CodeItem code)
        {
            if (code == null) return null;
            if (string.IsNullOrWhiteSpace(PdfFolder))
                InitializePdfFolder(null);

            try
            {
                // prefer Number
                if (!string.IsNullOrWhiteSpace(code.Number))
                {
                    var p = Path.Combine(PdfFolder, MakeSafeFileName(code.Number) + ".pdf");
                    if (File.Exists(p)) return p;
                }

                // then LatestCode
                if (!string.IsNullOrWhiteSpace(code.LatestCode))
                {
                    var p2 = Path.Combine(PdfFolder, MakeSafeFileName(code.LatestCode) + ".pdf");
                    if (File.Exists(p2)) return p2;
                }

                // then filename from Link
                if (!string.IsNullOrWhiteSpace(code.Link))
                {
                    try
                    {
                        if (Uri.IsWellFormedUriString(code.Link, UriKind.Absolute))
                        {
                            var uri = new Uri(code.Link);
                            var name = Path.GetFileName(uri.LocalPath);
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                var p3 = Path.Combine(PdfFolder, name);
                                if (File.Exists(p3)) return p3;
                            }
                        }
                        else
                        {
                            // treat Link as relative filename in folder
                            var rel = Path.Combine(PdfFolder, code.Link);
                            if (File.Exists(rel)) return rel;
                        }
                    }
                    catch
                    {
                        // ignore URI/IO parse errors
                    }
                }

                // fallback: look for any file whose filename (without ext) matches a candidate (case-insensitive)
                var candidates = new[] { code.Number, code.LatestCode }
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => MakeSafeFileName(s!))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (candidates.Count > 0 && Directory.Exists(PdfFolder))
                {
                    foreach (var f in Directory.GetFiles(PdfFolder, "*.pdf"))
                    {
                        var nameNoExt = Path.GetFileNameWithoutExtension(f);
                        if (candidates.Contains(nameNoExt))
                            return f;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        // -------------------------
        // Helpers (local to TabViewModel to avoid coupling)
        // -------------------------
        private static string MakeSafeFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Unnamed";

            var invalid = Path.GetInvalidFileNameChars();
            var chars = input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
            var safe = new string(chars);

            safe = Regex.Replace(safe, @"\s{2,}", " ").Trim();
            safe = Regex.Replace(safe, @"[\. ]+$", "");

            if (string.IsNullOrEmpty(safe))
                return "Unnamed";

            return safe;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}