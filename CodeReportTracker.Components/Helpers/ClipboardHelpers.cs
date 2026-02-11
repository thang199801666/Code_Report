using CodeReportTracker.Core.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Linq;

namespace CodeReportTracker.Components.Helpers
{
    internal static class ClipboardHelpers
    {
        public static readonly Dictionary<string, string> HeaderMap = new(StringComparer.OrdinalIgnoreCase)
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

        public static bool TryGetStringFromRaw(object raw, out string result)
        {
            result = null;
            if (raw == null) return false;
            try
            {
                if (raw is string s) { result = s; return true; }
                if (raw is Stream st)
                {
                    try { if (st.CanSeek) st.Position = 0; } catch { }
                    using var sr = new StreamReader(st, detectEncodingFromByteOrderMarks: true);
                    result = sr.ReadToEnd();
                    return true;
                }
                if (raw is byte[] ba)
                {
                    if (ba.Length >= 3 && ba[0] == 0xEF && ba[1] == 0xBB && ba[2] == 0xBF)
                        result = Encoding.UTF8.GetString(ba);
                    else
                        result = Encoding.UTF8.GetString(ba);
                    return true;
                }
                result = raw.ToString();
                return !string.IsNullOrEmpty(result);
            }
            catch
            {
                result = null;
                return false;
            }
        }

        public static string ExtractFragmentFromCfHtml(string cfHtml)
        {
            if (string.IsNullOrWhiteSpace(cfHtml)) return cfHtml;

            var startMatch = Regex.Match(cfHtml, @"StartFragment:(\d+)", RegexOptions.IgnoreCase);
            var endMatch = Regex.Match(cfHtml, @"EndFragment:(\d+)", RegexOptions.IgnoreCase);
            if (startMatch.Success && endMatch.Success &&
                int.TryParse(startMatch.Groups[1].Value, out var sPos) &&
                int.TryParse(endMatch.Groups[1].Value, out var ePos) &&
                sPos >= 0 && ePos > sPos && ePos <= cfHtml.Length)
            {
                try { return cfHtml.Substring(sPos, ePos - sPos); } catch { }
            }

            const string sm = "<!--StartFragment-->";
            const string em = "<!--EndFragment-->";
            var si = cfHtml.IndexOf(sm, StringComparison.OrdinalIgnoreCase);
            var ei = cfHtml.IndexOf(em, StringComparison.OrdinalIgnoreCase);
            if (si >= 0 && ei > si) return cfHtml.Substring(si + sm.Length, ei - (si + sm.Length));

            var tableMatch = Regex.Match(cfHtml, @"<table[\s\S]*?</table>", RegexOptions.IgnoreCase);
            if (tableMatch.Success) return tableMatch.Value;

            return cfHtml;
        }

        public static void ParseHtmlTable(string html, out List<string> rowsText, out List<string[]> rowsLinks)
        {
            rowsText = new List<string>();
            rowsLinks = new List<string[]>();
            if (string.IsNullOrWhiteSpace(html)) return;

            try
            {
                var rowMatches = Regex.Matches(html, @"<tr[\s\S]*?>[\s\S]*?<\/tr\s*>", RegexOptions.IgnoreCase);
                if (rowMatches.Count == 0)
                {
                    var fallbackRows = Regex.Split(html, @"</tr\s*>", RegexOptions.IgnoreCase);
                    foreach (var r in fallbackRows)
                    {
                        if (string.IsNullOrWhiteSpace(r)) continue;
                        var cells = Regex.Matches(r, @"<(td|th)[\s\S]*?>[\s\S]*?<\/\1\s*>", RegexOptions.IgnoreCase);
                        if (cells.Count == 0) continue;
                        var texts = new List<string>();
                        var links = new List<string>();
                        foreach (Match c in cells)
                        {
                            var inner = c.Value;
                            var ah = Regex.Match(inner, @"<a\b[^>]*?\bhref\s*=\s*[""'](?<href>[^""']+)[""'][^>]*>(?<text>[\s\S]*?)<\/a\s*>", RegexOptions.IgnoreCase);
                            if (ah.Success)
                            {
                                var href = WebUtility.HtmlDecode(ah.Groups["href"].Value).Trim();
                                var anchorText = Regex.Replace(ah.Groups["text"].Value, @"<[^>]+>", string.Empty);
                                texts.Add(WebUtility.HtmlDecode(anchorText).Trim());
                                links.Add(href);
                            }
                            else
                            {
                                var cellText = Regex.Replace(inner, @"<[^>]+>", string.Empty);
                                texts.Add(WebUtility.HtmlDecode(cellText).Trim());
                                links.Add(null);
                            }
                        }
                        rowsText.Add(string.Join("\t", texts));
                        rowsLinks.Add(links.ToArray());
                    }
                }
                else
                {
                    foreach (Match rm in rowMatches)
                    {
                        var rowHtml = rm.Value;
                        var cellMatches = Regex.Matches(rowHtml, @"<(td|th)[\s\S]*?>[\s\S]*?<\/\1\s*>", RegexOptions.IgnoreCase);
                        var texts = new List<string>();
                        var links = new List<string>();
                        foreach (Match cm in cellMatches)
                        {
                            var cellHtml = cm.Value;
                            var ah = Regex.Match(cellHtml, @"<a\b[^>]*?\bhref\s*=\s*[""'](?<href>[^""']+)[""'][^>]*>(?<text>[\s\S]*?)<\/a\s*>", RegexOptions.IgnoreCase);
                            if (ah.Success)
                            {
                                var href = WebUtility.HtmlDecode(ah.Groups["href"].Value).Trim();
                                var anchorText = Regex.Replace(ah.Groups["text"].Value, @"<[^>]+>", string.Empty);
                                texts.Add(WebUtility.HtmlDecode(anchorText).Trim());
                                links.Add(href);
                            }
                            else
                            {
                                var cellText = Regex.Replace(cellHtml, @"<[^>]+>", string.Empty);
                                texts.Add(WebUtility.HtmlDecode(cellText).Trim());
                                links.Add(null);
                            }
                        }
                        if (texts.Count > 0)
                        {
                            rowsText.Add(string.Join("\t", texts));
                            rowsLinks.Add(links.ToArray());
                        }
                    }
                }
            }
            catch
            {
                rowsText = new List<string>();
                rowsLinks = new List<string[]>();
            }
        }

        public static string ConvertHtmlTableToText(string html)
        {
            try
            {
                const string startMarker = "<!--StartFragment-->";
                const string endMarker = "<!--EndFragment-->";
                string fragment = null;
                var si = html.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase);
                var ei = html.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase);
                if (si >= 0 && ei > si) fragment = html.Substring(si + startMarker.Length, ei - (si + startMarker.Length));
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

        public static string ConvertExcelXmlToText(string xml)
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
                        if (dataDirect.Length > 0) { outputLines.Add(string.Join("\t", dataDirect)); continue; }

                        var text = CleanCellText(row.Value ?? string.Empty);
                        if (!string.IsNullOrEmpty(text)) outputLines.Add(text);
                        continue;
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
            catch { return string.Empty; }
        }

        public static string NormalizeCsvToTabText(string csv)
        {
            var lines = csv.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (!line.Contains("\"")) { lines[i] = line.Replace(',', '\t'); continue; }

                var fields = new List<string>();
                bool inQuotes = false;
                var current = new StringBuilder();
                for (int j = 0; j < line.Length; j++)
                {
                    var ch = line[j];
                    if (ch == '"')
                    {
                        if (inQuotes && j + 1 < line.Length && line[j + 1] == '"') { current.Append('"'); j++; }
                        else inQuotes = !inQuotes;
                    }
                    else if (ch == ',' && !inQuotes) { fields.Add(current.ToString()); current.Clear(); }
                    else current.Append(ch);
                }
                fields.Add(current.ToString());
                lines[i] = string.Join("\t", fields.Select(f => CleanCellText(f.Trim().Trim('"'))));
            }
            return string.Join("\r\n", lines);
        }

        public static string CleanCellText(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = WebUtility.HtmlDecode(s).Trim();
            var sb = new StringBuilder(s.Length);
            foreach (var ch in s)
            {
                if (ch == '\t' || ch == '\r' || ch == '\n' ||
                    (!char.IsControl(ch) && ch != '\u200B' && ch != '\u200C' && ch != '\u200D' && ch != '\uFEFF' && ch != '\u2060'))
                {
                    sb.Append(ch);
                }
            }
            var collapsed = Regex.Replace(sb.ToString(), @"[ \t\u00A0]{2,}", " ");
            return collapsed.Trim();
        }

        public static bool HasProperty(Type type, string name)
        {
            return type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase) != null;
        }

        public static bool ParseBoolLike(string text)
        {
            var lower = text?.Trim().ToLowerInvariant() ?? string.Empty;
            if (string.IsNullOrEmpty(lower)) return false;
            if (lower == "1" || lower == "true" || lower == "yes" || lower == "y" || lower == "t") return true;
            if (lower == "0" || lower == "false" || lower == "no" || lower == "n" || lower == "f") return false;
            if (bool.TryParse(lower, out var b)) return b;
            return false;
        }

        public static object ConvertStringToType(string text, Type targetType)
        {
            if (string.IsNullOrEmpty(text))
            {
                if (Nullable.GetUnderlyingType(targetType) != null || !targetType.IsValueType) return null;
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
                if (nonNullableType == typeof(long)) { if (long.TryParse(text, out var l)) return l; return 0L; }
                if (nonNullableType == typeof(bool)) return ParseBoolLike(text);
                if (nonNullableType == typeof(DateTime)) { if (DateTime.TryParse(text, out var dt)) return dt; return DateTime.MinValue; }
                if (nonNullableType.IsEnum) { try { return Enum.Parse(nonNullableType, text, true); } catch { return Activator.CreateInstance(nonNullableType); } }
                return Convert.ChangeType(text, nonNullableType, CultureInfo.InvariantCulture);
            }
            catch { return nonNullableType.IsValueType ? Activator.CreateInstance(nonNullableType) : null; }
        }

        // Accept optional providedColumnNames: when caller (UI) knows DataGrid column bindings,
        // pass them so helper can map by position when there is no header row.
        public static List<(CodeItem Item, HashSet<string> SetProps)> ParseClipboardToCodeItemsWithSetProps(IDataObject dobj, IEnumerable<string> providedColumnNames = null)
        {
            var result = new List<(CodeItem, HashSet<string>)>();
            if (dobj == null) return result;

            string text = null;
            string rawHtml = null;
            List<string> htmlRowsText = null;
            List<string[]> htmlRowsLinks = null;

            try
            {
                // 1) Prefer CF_HTML
                if (dobj.GetDataPresent(DataFormats.Html))
                {
                    var raw = dobj.GetData(DataFormats.Html);
                    if (TryGetStringFromRaw(raw, out var htmlRaw) && !string.IsNullOrWhiteSpace(htmlRaw))
                    {
                        rawHtml = htmlRaw;
                        var fragment = ExtractFragmentFromCfHtml(htmlRaw);
                        ParseHtmlTable(fragment, out htmlRowsText, out htmlRowsLinks);

                        var conv = ConvertHtmlTableToText(fragment);
                        if (!string.IsNullOrWhiteSpace(conv))
                            text = conv;
                        else if (htmlRowsText != null && htmlRowsText.Count > 0)
                            text = string.Join("\r\n", htmlRowsText);
                    }
                }

                // 2) Spreadsheet/XML formats
                if (string.IsNullOrWhiteSpace(text))
                {
                    string[] formats;
                    try { formats = dobj.GetFormats(); } catch { formats = Array.Empty<string>(); }

                    var chosen = formats.FirstOrDefault(f => !string.IsNullOrWhiteSpace(f) &&
                                                             (f.IndexOf("xml", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                              f.IndexOf("spreadsheet", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                              f.IndexOf("excel", StringComparison.OrdinalIgnoreCase) >= 0));
                    if (!string.IsNullOrWhiteSpace(chosen))
                    {
                        var raw = dobj.GetData(chosen);
                        if (TryGetStringFromRaw(raw, out var xmlString) && !string.IsNullOrWhiteSpace(xmlString))
                        {
                            var conv = ConvertExcelXmlToText(xmlString);
                            if (!string.IsNullOrWhiteSpace(conv)) text = conv;
                        }
                    }
                }

                // 3) Plain text fallback
                if (string.IsNullOrWhiteSpace(text))
                {
                    if (dobj.GetDataPresent(DataFormats.UnicodeText))
                        text = dobj.GetData(DataFormats.UnicodeText) as string;
                    else if (dobj.GetDataPresent(DataFormats.Text))
                        text = dobj.GetData(DataFormats.Text) as string;
                }
            }
            catch
            {
                try
                {
                    if (dobj.GetDataPresent(DataFormats.UnicodeText))
                        text = dobj.GetData(DataFormats.UnicodeText) as string;
                    else if (dobj.GetDataPresent(DataFormats.Text))
                        text = dobj.GetData(DataFormats.Text) as string;
                }
                catch { text = null; }
            }

            if (string.IsNullOrWhiteSpace(text)) return result;

            if (!text.Contains("\t") && text.Contains(","))
                text = NormalizeCsvToTabText(text);

            // Extract HYPERLINK(...) formulas
            List<string[]> formulaLinks = null;
            if (text.IndexOf("HYPERLINK", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var rawLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                var processed = new List<string>(rawLines.Length);
                formulaLinks = new List<string[]>();
                foreach (var ln in rawLines)
                {
                    var parts = ln.Split('\t');
                    var linkRow = new string[parts.Length];
                    for (int i = 0; i < parts.Length; i++)
                    {
                        var p = parts[i];
                        var m = Regex.Match(p, @"=HYPERLINK\(\s*['""](?<url>[^'""]+)['""]\s*,\s*['""](?<text>[^'""]+)['""]\s*\)", RegexOptions.IgnoreCase);
                        if (m.Success)
                        {
                            linkRow[i] = WebUtility.HtmlDecode(m.Groups["url"].Value).Trim();
                            parts[i] = m.Groups["text"].Value;
                        }
                    }
                    formulaLinks.Add(linkRow);
                    processed.Add(string.Join("\t", parts));
                }
                text = string.Join("\r\n", processed);
            }

            // Build rows + per-cell links
            string[] rows;
            string[][] perCellLinks = null;
            if (htmlRowsText != null && htmlRowsText.Count > 0)
            {
                rows = htmlRowsText.ToArray();
                perCellLinks = htmlRowsLinks?.Select(a => a).ToArray();
            }
            else
            {
                rows = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (formulaLinks != null && formulaLinks.Count > 0)
                    perCellLinks = formulaLinks.Select(a => a).ToArray();
            }

            if (rows.Length == 0) return result;

            // Header detection (same heuristic as before)
            var firstCells = rows[0].Split('\t');
            var cleanedFirstCells = firstCells.Select(c => CleanCellText(c)).ToArray();

            int headerMatches = cleanedFirstCells.Count(c =>
                !string.IsNullOrWhiteSpace(c) &&
                (HeaderMap.ContainsKey(c.Trim()) || HasProperty(typeof(CodeItem), c.Trim()) ||
                 Regex.IsMatch(c, @"^[A-Za-z\s\-_]+$")));

            bool firstRowIsHeader = headerMatches >= Math.Max(2, (int)Math.Ceiling(cleanedFirstCells.Length * 0.5));
            int dataStart = firstRowIsHeader ? 1 : 0;

            // Resolve column names. If caller provided column names (DataGrid bindings), use them when there's no header.
            var columnNames = new List<string>();
            if (firstRowIsHeader)
            {
                for (int i = 0; i < cleanedFirstCells.Length; i++)
                {
                    var rawHeader = cleanedFirstCells[i];
                    if (string.IsNullOrWhiteSpace(rawHeader)) { columnNames.Add(string.Empty); continue; }

                    if (HeaderMap.TryGetValue(rawHeader.Trim(), out var mapped)) { columnNames.Add(mapped); continue; }
                    if (HasProperty(typeof(CodeItem), rawHeader)) { columnNames.Add(rawHeader); continue; }

                    var norm = Regex.Replace(rawHeader, @"[^A-Za-z0-9_]", string.Empty);
                    if (!string.IsNullOrWhiteSpace(norm))
                    {
                        var prop = typeof(CodeItem).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .FirstOrDefault(p => string.Equals(Regex.Replace(p.Name, @"[^A-Za-z0-9_]", string.Empty), norm, StringComparison.OrdinalIgnoreCase));
                        if (prop != null) { columnNames.Add(prop.Name); continue; }
                    }

                    columnNames.Add(rawHeader);
                }
            }
            else
            {
                if (providedColumnNames != null)
                {
                    var provided = providedColumnNames.ToList();
                    // ensure length matches number of columns in first data row
                    var targetCount = cleanedFirstCells.Length;
                    for (int i = 0; i < targetCount; i++)
                    {
                        if (i < provided.Count) columnNames.Add(provided[i] ?? string.Empty);
                        else columnNames.Add(string.Empty);
                    }
                }
                else
                {
                    // use empty placeholders sized to first row
                    for (int i = 0; i < cleanedFirstCells.Length; i++) columnNames.Add(string.Empty);
                }
            }

            // Map rows -> CodeItem with SetProps
            for (int r = dataStart; r < rows.Length; r++)
            {
                var cells = rows[r].Split('\t');
                var item = new CodeItem();
                var setProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // attach OpenLinkAction
                item.OpenLinkAction = (link) =>
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(link))
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo { FileName = link, UseShellExecute = true, Verb = "open" };
                            System.Diagnostics.Process.Start(psi);
                        }
                    }
                    catch { }
                };

                for (int c = 0; c < cells.Length; c++)
                {
                    if (c >= columnNames.Count) break;
                    var mappedColName = columnNames[c];
                    if (string.IsNullOrWhiteSpace(mappedColName) && c == 0) mappedColName = "Number";
                    if (string.IsNullOrWhiteSpace(mappedColName)) continue;

                    string cellLink = null;
                    try
                    {
                        if (perCellLinks != null && (r - dataStart) < perCellLinks.Length)
                        {
                            var rowLinks = perCellLinks[r - dataStart];
                            if (rowLinks != null && c < rowLinks.Length) cellLink = rowLinks[c];
                        }
                    }
                    catch { cellLink = null; }

                    var raw = CleanCellText(cells[c] ?? string.Empty);

                    // extract HYPERLINK formula if present
                    if (string.IsNullOrWhiteSpace(cellLink))
                    {
                        var m = Regex.Match(raw, @"=HYPERLINK\(\s*['""](?<url>[^'""]+)['""]\s*,\s*['""](?<text>[^'""]+)['""]\s*\)", RegexOptions.IgnoreCase);
                        if (m.Success)
                        {
                            cellLink = WebUtility.HtmlDecode(m.Groups["url"].Value).Trim();
                            raw = m.Groups["text"].Value;
                        }
                    }

                    try
                    {
                        if (c == 0)
                        {
                            item.Number = raw; setProps.Add(nameof(item.Number));
                            if (!string.IsNullOrWhiteSpace(cellLink)) { item.Link = cellLink; setProps.Add(nameof(item.Link)); }
                            if (string.Equals(mappedColName.Trim(), "Link", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(item.Link))
                            {
                                item.Link = raw; setProps.Add(nameof(item.Link));
                            }
                            continue;
                        }

                        switch (mappedColName.Trim())
                        {
                            case "Number":
                                item.Number = raw; setProps.Add(nameof(item.Number));
                                if (!string.IsNullOrWhiteSpace(cellLink)) { item.Link = cellLink; setProps.Add(nameof(item.Link)); }
                                break;
                            case "Link":
                                item.Link = raw; setProps.Add(nameof(item.Link));
                                break;
                            case "LatestCode":
                                item.LatestCode = raw; setProps.Add(nameof(item.LatestCode));
                                break;
                            case "IssueDate":
                                item.IssueDate = raw; setProps.Add(nameof(item.IssueDate));
                                break;
                            case "ExpirationDate":
                                item.ExpirationDate = raw; setProps.Add(nameof(item.ExpirationDate));
                                break;
                            case "DownloadProcess":
                                if (int.TryParse(raw, out var dp)) item.DownloadProcess = dp;
                                else if (raw.EndsWith("%") && int.TryParse(raw.TrimEnd('%').Trim(), out dp)) item.DownloadProcess = dp;
                                setProps.Add(nameof(item.DownloadProcess));
                                break;
                            case "LastCheck":
                                item.LastCheck = raw; setProps.Add(nameof(item.LastCheck));
                                break;
                            case "HasCheck":
                                item.HasCheck = ParseBoolLike(raw); setProps.Add(nameof(item.HasCheck));
                                break;
                            case "HasUpdate":
                                item.HasUpdate = ParseBoolLike(raw); setProps.Add(nameof(item.HasUpdate));
                                break;
                            default:
                                var prop = typeof(CodeItem).GetProperty(mappedColName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                                if (prop != null && prop.CanWrite)
                                {
                                    object parsed = ConvertStringToType(raw, prop.PropertyType);
                                    prop.SetValue(item, parsed);
                                    setProps.Add(prop.Name);
                                }
                                break;
                        }
                    }
                    catch { /* ignore per-cell failures */ }
                }

                result.Add((item, setProps));
            }

            return result;
        }
    }
}