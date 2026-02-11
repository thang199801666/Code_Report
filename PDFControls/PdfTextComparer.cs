using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using CodeReportTracker.Core.Models;

namespace PDFControls
{
    /// <summary>
    /// Helpers to parse first-page text of ICC-ES / IAPMO style PDFs and compare extracted fields.
    /// Mirrors the logic used previously in PDFReader.iapmoCheck / iccesCheck but returns structured results
    /// and comparison info so other parts of the app can consume differences programmatically.
    /// </summary>
    public static class PdfTextComparer
    {
        public class PdfCodeInfo
        {
            public string LatestCode { get; set; }
            public string IssueDate { get; set; }
            public string ExpirationDate { get; set; }
            public string RawText { get; set; }

            public PdfCodeInfo()
            {
                LatestCode = "n/a";
                IssueDate = "n/a";
                ExpirationDate = "n/a";
                RawText = string.Empty;
            }
        }

        public class PdfCompareResult
        {
            public PdfCodeInfo Left { get; set; }
            public PdfCodeInfo Right { get; set; }
            public List<string> Differences { get; set; }

            public PdfCompareResult()
            {
                Left = new PdfCodeInfo();
                Right = new PdfCodeInfo();
                Differences = new List<string>();
            }

            public bool HasDifferences
            {
                get { return Differences != null && Differences.Count > 0; }
            }
        }

        /// <summary>
        /// Parse first-page text using IAPMO-style heuristics.
        /// </summary>
        public static PdfCodeInfo ParseIapmo(string firstPageText)
        {
            var info = new PdfCodeInfo();
            if (string.IsNullOrEmpty(firstPageText))
            {
                info.RawText = string.Empty;
                return info;
            }

            info.RawText = firstPageText;

            // Patterns adapted from original code
            var latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
            var issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
            var revisedDateRegex = new Regex(@"(?s)(revised.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);

            var latestMatch = latestCodeRex.Match(firstPageText);
            var issueMatch = issueDateRegex.Match(firstPageText);
            var revisedMatch = revisedDateRegex.Match(firstPageText);

            // Latest code (4 digits)
            if (latestMatch != null && latestMatch.Groups.Count > 1)
            {
                var val = latestMatch.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(val))
                    info.LatestCode = val;
            }

            // Parse dates
            DateTime issueDate = DateTime.MinValue;
            DateTime revisedDate = DateTime.MinValue;
            DateTime expirationDate = DateTime.MinValue;

            if (issueMatch != null && issueMatch.Groups.Count > 2)
            {
                DateTime.TryParse(issueMatch.Groups[2].Value.Replace(" ", ""), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out issueDate);
            }

            if (revisedMatch != null && revisedMatch.Groups.Count > 2)
            {
                DateTime.TryParse(revisedMatch.Groups[2].Value.Replace(" ", ""), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out revisedDate);
            }

            // Try to find expiration/renewal/valid date more robustly
            var expirationKeywords = new[]
            {
                "renewal", "renewal date", "renew", "valid through", "valid thru", "valid until", "valid", "expiration", "expires", "expiration date"
            };
            DateTime found;
            if (TryFindDateAfterKeywords(firstPageText, expirationKeywords, out found))
            {
                expirationDate = found;
            }

            if (revisedDate > issueDate) issueDate = revisedDate;

            info.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
            info.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");

            return info;
        }

        /// <summary>
        /// Parse first-page text using ICC-ES-style heuristics.
        /// </summary>
        public static PdfCodeInfo ParseIccEs(string firstPageText)
        {
            var info = new PdfCodeInfo();
            if (string.IsNullOrEmpty(firstPageText))
            {
                info.RawText = string.Empty;
                return info;
            }

            info.RawText = firstPageText;

            var latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
            var issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
            var revisedDateRegex = new Regex(@"(?s)(Revised.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);

            var latestMatch = latestCodeRex.Match(firstPageText);
            var issueMatch = issueDateRegex.Match(firstPageText);
            var revisedMatch = revisedDateRegex.Match(firstPageText);

            if (latestMatch != null && latestMatch.Groups.Count > 1)
            {
                var val = latestMatch.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(val))
                    info.LatestCode = val;
            }

            DateTime issueDate = DateTime.MinValue;
            DateTime revisedDate = DateTime.MinValue;
            DateTime expirationDate = DateTime.MinValue;

            // Issue: combine groups 2 and 3 when present (some patterns split text + year)
            try
            {
                if (issueMatch != null && issueMatch.Groups.Count > 3)
                {
                    var candidate = (issueMatch.Groups[2].Value ?? string.Empty).Replace(" ", "") + (issueMatch.Groups[3].Value ?? string.Empty);
                    DateTime.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out issueDate);
                }
            }
            catch { }

            try
            {
                if (revisedMatch != null && revisedMatch.Groups.Count > 3)
                {
                    var candidate = (revisedMatch.Groups[2].Value ?? string.Empty).Replace(" ", "") + (revisedMatch.Groups[3].Value ?? string.Empty);
                    DateTime.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out revisedDate);
                }
            }
            catch { }

            // Try to find expiration/renewal/valid date more robustly
            var expirationKeywords = new[]
            {
                "renewal", "renewal date", "renew", "valid through", "valid thru", "valid until", "valid", "expiration", "expires", "expiration date"
            };
            DateTime found;
            if (TryFindDateAfterKeywords(firstPageText, expirationKeywords, out found))
            {
                expirationDate = found;
            }

            if (revisedDate > issueDate) issueDate = revisedDate;

            info.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
            info.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");

            return info;
        }

        /// <summary>
        /// Compare two PDFReader instances using IAPMO parsing heuristics.
        /// Caller is responsible for obtaining first-page text from the PDFReader (e.g. via ExtractFirstPageText()).
        /// </summary>
        public static PdfCompareResult CompareIapmo(string leftFirstPageText, string rightFirstPageText)
        {
            var left = ParseIapmo(leftFirstPageText);
            var right = ParseIapmo(rightFirstPageText);

            var result = new PdfCompareResult();
            result.Left = left;
            result.Right = right;

            // Compare fields and record differences
            if (!string.Equals(left.LatestCode ?? string.Empty, right.LatestCode ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("LatestCode differs: Left='{0}' Right='{1}'", left.LatestCode, right.LatestCode));

            if (!string.Equals(left.IssueDate ?? string.Empty, right.IssueDate ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("IssueDate differs: Left='{0}' Right='{1}'", left.IssueDate, right.IssueDate));

            if (!string.Equals(left.ExpirationDate ?? string.Empty, right.ExpirationDate ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("ExpirationDate differs: Left='{0}' Right='{1}'", left.ExpirationDate, right.ExpirationDate));

            return result;
        }

        /// <summary>
        /// Compare two PDFReader instances using ICC-ES parsing heuristics.
        /// </summary>
        public static PdfCompareResult CompareIccEs(string leftFirstPageText, string rightFirstPageText)
        {
            var left = ParseIccEs(leftFirstPageText);
            var right = ParseIccEs(rightFirstPageText);

            var result = new PdfCompareResult();
            result.Left = left;
            result.Right = right;

            if (!string.Equals(left.LatestCode ?? string.Empty, right.LatestCode ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("LatestCode differs: Left='{0}' Right='{1}'", left.LatestCode, right.LatestCode));

            if (!string.Equals(left.IssueDate ?? string.Empty, right.IssueDate ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("IssueDate differs: Left='{0}' Right='{1}'", left.IssueDate, right.IssueDate));

            if (!string.Equals(left.ExpirationDate ?? string.Empty, right.ExpirationDate ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                result.Differences.Add(string.Format("ExpirationDate differs: Left='{0}' Right='{1}'", left.ExpirationDate, right.ExpirationDate));

            return result;
        }

        /// <summary>
        /// Equivalent of PDFReader.iapmoCheck but for CodeItem (view-model).
        /// Preserves old values on the supplied CodeItem, updates fields and flags.
        /// </summary>
        /// <param name="item">CodeItem to update (preserves _Old properties before overwriting)</param>
        /// <param name="firstPageText">First-page text obtained from a PDFReader instance (ExtractFirstPageText)</param>
        public static void IapmoCheck(CodeItem item, string firstPageText)
        {
            if (item == null) throw new ArgumentNullException(nameof(item));

            var parsed = ParseIapmo(firstPageText ?? string.Empty);

            // Preserve previous values
            item.LatestCode_Old = item.LatestCode;
            item.IssueDate_Old = item.IssueDate;
            item.ExpirationDate_Old = item.ExpirationDate;

            // Normalize values and assign
            item.LatestCode = string.IsNullOrEmpty(parsed.LatestCode) ? "n/a" : parsed.LatestCode;
            item.IssueDate = string.IsNullOrEmpty(parsed.IssueDate) ? "n/a" : parsed.IssueDate;
            item.ExpirationDate = string.IsNullOrEmpty(parsed.ExpirationDate) ? "n/a" : parsed.ExpirationDate;

            // Determine whether meaningful changes occurred (case-insensitive)
            var hasUpdate = false;
            if (!string.Equals(item.LatestCode ?? string.Empty, item.LatestCode_Old ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                hasUpdate = true;
            if (!string.Equals(item.IssueDate ?? string.Empty, item.IssueDate_Old ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                hasUpdate = true;
            if (!string.Equals(item.ExpirationDate ?? string.Empty, item.ExpirationDate_Old ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                hasUpdate = true;

            item.HasCheck = true;
            item.HasUpdate = hasUpdate;
            item.LastCheck = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        // --- Helper: robust extraction of dates after keywords ---
        private static bool TryFindDateAfterKeywords(string text, string[] keywords, out DateTime found)
        {
            found = DateTime.MinValue;
            if (string.IsNullOrEmpty(text)) return false;

            // Candidate date regexes (common formats)
            var datePatterns = new[]
            {
                new Regex(@"\b\d{1,2}/\d{1,2}/\d{2,4}\b", RegexOptions.IgnoreCase),
                new Regex(@"\b\d{1,2}-\d{1,2}-\d{2,4}\b", RegexOptions.IgnoreCase),
                new Regex(@"\b[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}\b", RegexOptions.IgnoreCase), // e.g. January 1, 2023
                new Regex(@"\b[A-Za-z]{3,9}\s+\d{4}\b", RegexOptions.IgnoreCase), // e.g. January 2023
                new Regex(@"\b\d{4}\b", RegexOptions.IgnoreCase) // fallback year
            };

            foreach (var kw in keywords)
            {
                var idx = IndexOfIgnoreCase(text, kw);
                if (idx < 0) continue;

                // examine a window after the keyword
                var windowStart = idx;
                var windowLen = Math.Min(300, text.Length - windowStart);
                var window = text.Substring(windowStart, windowLen);

                foreach (var rex in datePatterns)
                {
                    var m = rex.Match(window);
                    if (m.Success)
                    {
                        var candidate = m.Value.Trim();
                        if (DateTime.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out found))
                            return true;

                        // Try a more permissive parse without invariant culture
                        if (DateTime.TryParse(candidate, out found))
                            return true;

                        // If pattern was a 4-digit year, try to attach a month if present earlier in the window
                        if (regexFindMonthBeforeYear(window, m.Index, out var monthCandidate))
                        {
                            var combined = monthCandidate + " " + candidate;
                            if (DateTime.TryParse(combined, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out found))
                                return true;
                            if (DateTime.TryParse(combined, out found))
                                return true;
                        }
                    }
                }
            }

            // fallback: search whole text for any date pattern
            foreach (var rex in new[]
                     {
                         new Regex(@"\b\d{1,2}/\d{1,2}/\d{2,4}\b", RegexOptions.IgnoreCase),
                         new Regex(@"\b[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}\b", RegexOptions.IgnoreCase),
                         new Regex(@"\b[A-Za-z]{3,9}\s+\d{4}\b", RegexOptions.IgnoreCase),
                         new Regex(@"\b\d{4}\b", RegexOptions.IgnoreCase)
                     })
            {
                var m = rex.Match(text);
                if (m.Success)
                {
                    var candidate = m.Value.Trim();
                    if (DateTime.TryParse(candidate, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out found))
                        return true;
                    if (DateTime.TryParse(candidate, out found))
                        return true;
                }
            }

            return false;

            static int IndexOfIgnoreCase(string src, string value)
            {
                return src?.IndexOf(value ?? string.Empty, StringComparison.OrdinalIgnoreCase) ?? -1;
            }

            static bool regexFindMonthBeforeYear(string window, int yearIndex, out string month)
            {
                month = null;
                // look back up to 20 characters for a month name (e.g. "Jan", "January")
                var lookBackStart = Math.Max(0, yearIndex - 30);
                var segment = window.Substring(lookBackStart, Math.Min(30 + 10, window.Length - lookBackStart));
                var monthRex = new Regex(@"\b(Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)\b", RegexOptions.IgnoreCase);
                var m = monthRex.Match(segment);
                if (m.Success)
                {
                    month = m.Value;
                    return true;
                }

                return false;
            }
        }
    }
}