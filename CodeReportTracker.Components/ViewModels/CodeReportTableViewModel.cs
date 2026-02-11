using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Net;
using System.Xml.Linq;
using CodeReportTracker.Core.Models;

namespace CodeReportTracker.Components.ViewModels
{
    public sealed class CodeReportTableViewModel : System.ComponentModel.INotifyPropertyChanged
    {
        private CodeItem? _selectedItem;

        public ObservableCollection<CodeItem> Items { get; } = new();

        public CodeItem? SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (ReferenceEquals(_selectedItem, value)) return;
                _selectedItem = value;
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(SelectedItem)));
            }
        }

        public ICommand PasteCommand { get; }
        public ICommand ImportCommand { get; }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

        public CodeReportTableViewModel()
        {
            PasteCommand = new RelayCommand(_ => PasteFromClipboard());
            ImportCommand = new RelayCommand(_ => { /* placeholder for import logic */ });
        }

        // Exposed paste behavior moved into VM so UI (XAML) can bind command in MVVM style.
        private void PasteFromClipboard()
        {
            IDataObject dobj;
            try
            {
                dobj = Clipboard.GetDataObject();
            }
            catch
            {
                return;
            }

            if (dobj == null) return;

            string? text = null;

            try
            {
                var formats = dobj.GetFormats();
                var xmlFormat = formats.FirstOrDefault(f => !string.IsNullOrWhiteSpace(f) &&
                                                            (f.IndexOf("xml", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                             f.IndexOf("spreadsheet", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                             f.IndexOf("excel", StringComparison.OrdinalIgnoreCase) >= 0));
                if (!string.IsNullOrEmpty(xmlFormat))
                {
                    var raw = dobj.GetData(xmlFormat);
                    if (raw != null)
                    {
                        string? xmlString = null;
                        XDocument? xmlDoc = null;

                        if (raw is string s) xmlString = s;
                        else if (raw is Stream st)
                        {
                            try
                            {
                                st.Position = 0;
                                using var sr = new StreamReader(st, detectEncodingFromByteOrderMarks: true);
                                xmlString = sr.ReadToEnd();
                            }
                            catch { xmlString = null; }
                        }
                        else if (raw is MemoryStream ms)
                        {
                            try
                            {
                                ms.Position = 0;
                                using var sr = new StreamReader(ms, detectEncodingFromByteOrderMarks: true);
                                xmlString = sr.ReadToEnd();
                            }
                            catch { xmlString = null; }
                        }
                        else if (raw is byte[] ba)
                        {
                            try
                            {
                                if (ba.Length >= 3 && ba[0] == 0xEF && ba[1] == 0xBB && ba[2] == 0xBF)
                                    xmlString = Encoding.UTF8.GetString(ba);
                                else if (ba.Length >= 2 && ba[0] == 0xFF && ba[1] == 0xFE)
                                    xmlString = Encoding.Unicode.GetString(ba);
                                else if (ba.Length >= 2 && ba[0] == 0xFE && ba[1] == 0xFF)
                                    xmlString = Encoding.BigEndianUnicode.GetString(ba);
                                else
                                    xmlString = Encoding.UTF8.GetString(ba);
                            }
                            catch { xmlString = null; }
                        }
                        else
                        {
                            try { xmlString = raw.ToString(); } catch { xmlString = null; }
                        }

                        if (!string.IsNullOrWhiteSpace(xmlString))
                        {
                            var converted = ConvertExcelXmlToText(xmlString);
                            if (!string.IsNullOrWhiteSpace(converted))
                                text = converted;
                        }

                        if (string.IsNullOrWhiteSpace(text))
                        {
                            try
                            {
                                if (raw is Stream streamForDoc)
                                {
                                    streamForDoc.Position = 0;
                                    xmlDoc = XDocument.Load(streamForDoc);
                                }
                                else if (!string.IsNullOrWhiteSpace(xmlString))
                                {
                                    xmlDoc = XDocument.Parse(xmlString);
                                }
                            }
                            catch { xmlDoc = null; }

                            if (xmlDoc != null)
                            {
                                List<string> ExtractRowsFromSpreadsheetXml(XDocument doc)
                                {
                                    try
                                    {
                                        var rows = doc.Descendants().Where(x => string.Equals(x.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase));
                                        var outLines = new List<string>();

                                        foreach (var row in rows)
                                        {
                                            var cellElements = row.Elements().Where(e => string.Equals(e.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase)).ToList();
                                            if (cellElements.Count == 0)
                                            {
                                                var dataDirect = row.Elements().Where(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase))
                                                                              .Select(d => CleanCellText(d.Value ?? string.Empty)).ToArray();
                                                if (dataDirect.Length > 0)
                                                {
                                                    outLines.Add(string.Join("\t", dataDirect));
                                                    continue;
                                                }
                                                else
                                                {
                                                    var textFallback = CleanCellText(row.Value ?? string.Empty);
                                                    if (!string.IsNullOrEmpty(textFallback)) outLines.Add(textFallback);
                                                    continue;
                                                }
                                            }

                                            var values = new List<string>();
                                            int expectedIndex = 1;
                                            foreach (var cell in cellElements)
                                            {
                                                var indexAttr = cell.Attributes().FirstOrDefault(a => a.Name.LocalName.Equals("Index", StringComparison.OrdinalIgnoreCase));
                                                if (indexAttr != null && int.TryParse(indexAttr.Value, out var idx))
                                                {
                                                    while (expectedIndex < idx)
                                                    {
                                                        values.Add(string.Empty);
                                                        expectedIndex++;
                                                    }
                                                }

                                                string value = string.Empty;
                                                var data = cell.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase));
                                                if (data != null) value = data.Value ?? string.Empty;
                                                else value = cell.Value ?? string.Empty;

                                                values.Add(CleanCellText(value));
                                                expectedIndex++;
                                            }

                                            outLines.Add(string.Join("\t", values));
                                        }

                                        if (outLines.Count == 0)
                                        {
                                            var table = doc.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Table", StringComparison.OrdinalIgnoreCase));
                                            if (table != null)
                                            {
                                                foreach (var row in table.Elements().Where(e => string.Equals(e.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase)))
                                                {
                                                    var vals = new List<string>();
                                                    foreach (var cell in row.Elements().Where(e => string.Equals(e.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase)))
                                                    {
                                                        var data = cell.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase));
                                                        vals.Add(CleanCellText((data?.Value ?? string.Empty)));
                                                    }
                                                    outLines.Add(string.Join("\t", vals));
                                                }
                                            }
                                        }

                                        return outLines;
                                    }
                                    catch
                                    {
                                        return new List<string>();
                                    }
                                }

                                var extracted = ExtractRowsFromSpreadsheetXml(xmlDoc);
                                if (extracted.Count > 0)
                                    text = string.Join("\r\n", extracted);
                            }
                        }
                    }
                }
            }
            catch
            {
                // ignore xml detection failures and fall back to other formats
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                if (dobj.GetDataPresent(DataFormats.UnicodeText))
                {
                    text = dobj.GetData(DataFormats.UnicodeText) as string;
                }
                else if (dobj.GetDataPresent(DataFormats.Text))
                {
                    text = dobj.GetData(DataFormats.Text) as string;
                }
                else if (dobj.GetDataPresent(DataFormats.Html))
                {
                    var html = dobj.GetData(DataFormats.Html) as string;
                    if (!string.IsNullOrWhiteSpace(html))
                        text = ConvertHtmlTableToText(html);
                }
            }

            if (string.IsNullOrWhiteSpace(text)) return;

            if (!text.Contains("\t") && text.Contains(","))
            {
                text = NormalizeCsvToTabText(text);
            }

            var rows = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (rows.Length == 0) return;

            var firstCells = rows[0].Split('\t');
            bool firstRowIsHeader = firstCells.Any(c => !string.IsNullOrWhiteSpace(c) && c.Trim().Any(char.IsLetter));

            List<string> columnNames = new();
            int dataStart = 0;

            if (firstRowIsHeader)
            {
                columnNames = firstCells.Select(c => CleanCellText(c)).ToList();
                dataStart = 1;
            }
            else
            {
                // No header supplied — attempt to guess columns by order matching CodeItem properties in headerMap below.
            }

            var headerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Code Report No", "Number" },
                { "Number", "Number" },
                { "Link", "Link" },
                { "Latest Code", "LatestCode" },
                { "Issue/Rev Date", "IssueDate" },
                { "Expiration Date", "ExpirationDate" },
                { "Download Process", "DownloadProcess" },
                { "DownloadProcess", "DownloadProcess" },
                { "Last Check", "LastCheck" },
                { "Is Check", "HasCheck" },
                { "Is Update", "HasUpdate" },
                { "Status", "HasCheck" }
            };

            for (int r = dataStart; r < rows.Length; r++)
            {
                var cells = rows[r].Split('\t');
                var item = new CodeItem();
                var setProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                item.OpenLinkAction = (link) =>
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(link))
                        {
                            var psi = new ProcessStartInfo
                            {
                                FileName = link,
                                UseShellExecute = true,
                                Verb = "open"
                            };
                            Process.Start(psi);
                        }
                    }
                    catch { /* swallow */ }
                };

                for (int c = 0; c < cells.Length; c++)
                {
                    var colName = (firstRowIsHeader && c < columnNames.Count) ? columnNames[c]?.Trim() ?? string.Empty : string.Empty;
                    if (string.IsNullOrWhiteSpace(colName))
                    {
                        // If no header, attempt positional mapping (not implemented here) — skip unknowns.
                    }

                    if (!HasProperty(typeof(CodeItem), colName))
                    {
                        if (headerMap.TryGetValue(colName, out var mapped))
                            colName = mapped;
                    }

                    var raw = CleanCellText(cells[c] ?? string.Empty);

                    try
                    {
                        switch (colName.Trim())
                        {
                            case "Number":
                                item.Number = raw;
                                setProps.Add(nameof(item.Number));
                                break;
                            case "Link":
                                item.Link = raw;
                                setProps.Add(nameof(item.Link));
                                break;
                            case "LatestCode":
                                item.LatestCode = raw;
                                setProps.Add(nameof(item.LatestCode));
                                break;
                            case "IssueDate":
                                item.IssueDate = raw;
                                setProps.Add(nameof(item.IssueDate));
                                break;
                            case "ExpirationDate":
                                item.ExpirationDate = raw;
                                setProps.Add(nameof(item.ExpirationDate));
                                break;
                            case "DownloadProcess":
                                if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dp)) item.DownloadProcess = dp;
                                else if (raw.EndsWith("%") && int.TryParse(raw.TrimEnd('%').Trim(), out dp)) item.DownloadProcess = dp;
                                else if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) item.DownloadProcess = (int)d;
                                setProps.Add(nameof(item.DownloadProcess));
                                break;
                            case "LastCheck":
                                item.LastCheck = raw;
                                setProps.Add(nameof(item.LastCheck));
                                break;
                            case "HasCheck":
                                item.HasCheck = ParseBoolLike(raw);
                                setProps.Add(nameof(item.HasCheck));
                                break;
                            case "HasUpdate":
                                item.HasUpdate = ParseBoolLike(raw);
                                setProps.Add(nameof(item.HasUpdate));
                                break;
                            default:
                                var prop = typeof(CodeItem).GetProperty(colName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                                if (prop != null && prop.CanWrite)
                                {
                                    object parsed = ConvertStringToType(raw, prop.PropertyType);
                                    prop.SetValue(item, parsed);
                                    setProps.Add(prop.Name);
                                }
                                break;
                        }
                    }
                    catch { /* ignore per-cell errors */ }
                }

                // Try to find existing by Number then Link and update; otherwise add new.
                CodeItem? found = null;
                if (!string.IsNullOrWhiteSpace(item.Number))
                {
                    found = Items.FirstOrDefault(ci => !string.IsNullOrWhiteSpace(ci.Number) && string.Equals(ci.Number?.Trim(), item.Number?.Trim(), StringComparison.OrdinalIgnoreCase));
                }

                if (found == null && !string.IsNullOrWhiteSpace(item.Link))
                {
                    found = Items.FirstOrDefault(ci => !string.IsNullOrWhiteSpace(ci.Link) && string.Equals(ci.Link?.Trim(), item.Link?.Trim(), StringComparison.OrdinalIgnoreCase));
                }

                if (found != null)
                {
                    var targetType = found.GetType();
                    foreach (var propName in setProps)
                    {
                        var prop = targetType.GetProperty(propName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                        if (prop != null && prop.CanWrite)
                        {
                            var newValProp = typeof(CodeItem).GetProperty(prop.Name, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                            if (newValProp != null)
                            {
                                var value = newValProp.GetValue(item);
                                try { prop.SetValue(found, value); } catch { }
                            }
                        }
                    }
                }
                else
                {
                    Items.Add(item);
                }
            }
        }

        #region Helpers (copied-and-slimmed from code-behind)

        private static string ConvertExcelXmlToText(string xml)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(xml)) return string.Empty;
                XDocument doc = XDocument.Parse(xml);
                if (doc.Root == null) return string.Empty;
                var rows = doc.Descendants().Where(x => string.Equals(x.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase));
                var outputLines = new List<string>();
                foreach (var row in rows)
                {
                    var cells = row.Elements().Where(e => string.Equals(e.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase)).ToList();
                    if (cells.Count == 0)
                    {
                        var dataDirect = row.Elements().Where(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase))
                                                      .Select(d => CleanCellText(d.Value ?? string.Empty)).ToArray();
                        if (dataDirect.Length > 0)
                        {
                            outputLines.Add(string.Join("\t", dataDirect));
                            continue;
                        }
                        else
                        {
                            var text = CleanCellText(row.Value ?? string.Empty);
                            if (!string.IsNullOrEmpty(text)) outputLines.Add(text);
                            continue;
                        }
                    }

                    var cellValues = new List<string>();
                    foreach (var cell in cells)
                    {
                        var data = cell.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase));
                        var value = data != null ? data.Value ?? string.Empty : cell.Value ?? string.Empty;
                        cellValues.Add(CleanCellText(value));
                    }

                    outputLines.Add(string.Join("\t", cellValues));
                }

                if (outputLines.Count == 0)
                {
                    var table = doc.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Table", StringComparison.OrdinalIgnoreCase));
                    if (table != null)
                    {
                        foreach (var row in table.Elements().Where(e => string.Equals(e.Name.LocalName, "Row", StringComparison.OrdinalIgnoreCase)))
                        {
                            var cellValues = new List<string>();
                            foreach (var cell in row.Elements().Where(e => string.Equals(e.Name.LocalName, "Cell", StringComparison.OrdinalIgnoreCase)))
                            {
                                var data = cell.Elements().FirstOrDefault(e => string.Equals(e.Name.LocalName, "Data", StringComparison.OrdinalIgnoreCase));
                                cellValues.Add(CleanCellText((data?.Value ?? string.Empty)));
                            }
                            outputLines.Add(string.Join("\t", cellValues));
                        }
                    }
                }

                return string.Join("\r\n", outputLines.Where(l => l != null));
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ConvertHtmlTableToText(string html)
        {
            try
            {
                var startMarker = "<!--StartFragment-->";
                var endMarker = "<!--EndFragment-->";
                string? fragment = null;
                var si = html.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase);
                var ei = html.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase);
                if (si >= 0 && ei > si)
                {
                    fragment = html.Substring(si + startMarker.Length, ei - (si + startMarker.Length));
                }
                else
                {
                    var tableMatch = Regex.Match(html, @"<table[\s\S]*?</table>", RegexOptions.IgnoreCase);
                    fragment = tableMatch.Success ? tableMatch.Value : html;
                }

                fragment = Regex.Replace(fragment, @"</tr\s*>", "\n", RegexOptions.IgnoreCase);
                fragment = Regex.Replace(fragment, @"</t[dh]\s*>", "\t", RegexOptions.IgnoreCase);
                fragment = Regex.Replace(fragment, @"<[^>]+>", string.Empty);
                fragment = WebUtility.HtmlDecode(fragment).Replace("\r\n", "\n").Replace("\r", "\n");

                var lines = fragment.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(l => CleanCellText(l.TrimEnd('\t').Trim()))
                                    .ToArray();

                return string.Join("\r\n", lines);
            }
            catch { return string.Empty; }
        }

        private static string NormalizeCsvToTabText(string csv)
        {
            var lines = csv.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (!line.Contains("\""))
                {
                    lines[i] = line.Replace(',', '\t');
                    continue;
                }

                var fields = new List<string>();
                bool inQuotes = false;
                var current = "";
                for (int j = 0; j < line.Length; j++)
                {
                    var ch = line[j];
                    if (ch == '"')
                    {
                        if (inQuotes && j + 1 < line.Length && line[j + 1] == '"')
                        {
                            current += '"';
                            j++;
                        }
                        else
                        {
                            inQuotes = !inQuotes;
                        }
                    }
                    else if (ch == ',' && !inQuotes)
                    {
                        fields.Add(current);
                        current = "";
                    }
                    else
                    {
                        current += ch;
                    }
                }
                fields.Add(current);
                lines[i] = string.Join("\t", fields.Select(f => CleanCellText(f.Trim().Trim('"'))));
            }

            return string.Join("\r\n", lines);
        }

        private static string CleanCellText(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = WebUtility.HtmlDecode(s);
            s = s.Trim();
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (ch == '\t' || ch == '\r' || ch == '\n' ||
                    !char.IsControl(ch) &&
                    ch != '\u200B' && ch != '\u200C' && ch != '\u200D' && ch != '\uFEFF' &&
                    ch != '\u2060')
                {
                    sb.Append(ch);
                }
            }

            var collapsed = Regex.Replace(sb.ToString(), @"[ \t\u00A0]{2,}", " ");
            return collapsed.Trim();
        }

        private static bool HasProperty(Type type, string name)
        {
            return type.GetProperty(name, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase) != null;
        }

        private static bool ParseBoolLike(string text)
        {
            var lower = text?.Trim().ToLowerInvariant() ?? string.Empty;
            if (string.IsNullOrEmpty(lower)) return false;
            if (lower == "1" || lower == "true" || lower == "yes" || lower == "y" || lower == "t") return true;
            if (lower == "0" || lower == "false" || lower == "no" || lower == "n" || lower == "f") return false;
            if (bool.TryParse(lower, out var b)) return b;
            return false;
        }

        private static object? ConvertStringToType(string text, Type targetType)
        {
            if (string.IsNullOrEmpty(text))
            {
                if (Nullable.GetUnderlyingType(targetType) != null || !targetType.IsValueType)
                    return null;
                return Activator.CreateInstance(targetType);
            }

            var nonNullableType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (nonNullableType == typeof(string)) return text;
                if (nonNullableType == typeof(int))
                {
                    if (int.TryParse(text, out var i)) return i;
                    if (text.EndsWith("%") && int.TryParse(text.TrimEnd('%').Trim(), out i)) return i;
                    if (double.TryParse(text, out var d)) return (int)d;
                    return 0;
                }
                if (nonNullableType == typeof(long))
                {
                    if (long.TryParse(text, out var l)) return l;
                    return 0L;
                }
                if (nonNullableType == typeof(bool))
                {
                    return ParseBoolLike(text);
                }
                if (nonNullableType == typeof(DateTime))
                {
                    if (DateTime.TryParse(text, out var dt)) return dt;
                    return DateTime.MinValue;
                }
                if (nonNullableType.IsEnum)
                {
                    try { return Enum.Parse(nonNullableType, text, true); }
                    catch { return Activator.CreateInstance(nonNullableType); }
                }

                return Convert.ChangeType(text, nonNullableType, CultureInfo.InvariantCulture);
            }
            catch
            {
                return nonNullableType.IsValueType ? Activator.CreateInstance(nonNullableType) : null;
            }
        }

        #endregion

        #region Simple RelayCommand (internal)

        private sealed class RelayCommand : ICommand
        {
            private readonly Action<object?> _execute;
            private readonly Func<object?, bool>? _canExecute;

            public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
            public void Execute(object? parameter) => _execute(parameter);
#pragma warning disable CS0067
            public event EventHandler? CanExecuteChanged;
#pragma warning restore CS0067
        }

        #endregion
    }
}