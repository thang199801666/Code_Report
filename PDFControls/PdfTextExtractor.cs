using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;

namespace PDFControls
{
    /// <summary>
    /// Lightweight helper for extracting text from PDFs using PdfPig (UglyToad.PdfPig).
    /// Includes per-page extraction and HTML conversion helpers suitable for rendering in a web view.
    /// Now supports reading PDFs from HTTP/HTTPS URLs.
    /// </summary>
    public static class PdfTextExtractor
    {
        private static readonly HttpClient s_httpClient = new HttpClient();
        private const int MaxEmbeddedPdfDepth = 3;
        private const int MaxEmbeddedPdfCandidates = 20;

        /// <summary>
        /// Extracts all text from a PDF file on disk (all pages concatenated).
        /// </summary>
        /// <param name="filePath">Absolute or relative path to the PDF file.</param>
        /// <returns>Concatenated text of all pages. Empty string when no text found.</returns>
        public static string ExtractTextFromFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("PDF file not found.", filePath);

            try
            {
                var sb = new StringBuilder();
                using var document = PdfDocument.Open(filePath);
                foreach (var page in document.GetPages().OrderBy(p => p.Number))
                {
                    if (!string.IsNullOrEmpty(page.Text))
                    {
                        sb.AppendLine(page.Text);
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract text from PDF file.", ex);
            }
        }

        /// <summary>
        /// Extracts all text from PDF bytes (all pages concatenated).
        /// </summary>
        public static string ExtractTextFromBytes(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            try
            {
                using var ms = new MemoryStream(bytes);
                var sb = new StringBuilder();
                using var document = PdfDocument.Open(ms);
                foreach (var page in document.GetPages().OrderBy(p => p.Number))
                {
                    if (!string.IsNullOrEmpty(page.Text))
                    {
                        sb.AppendLine(page.Text);
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract text from PDF bytes.", ex);
            }
        }

        /// <summary>
        /// Returns the number of pages in the PDF file.
        /// </summary>
        public static int GetPageCount(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("PDF file not found.", filePath);

            try
            {
                using var document = PdfDocument.Open(filePath);
                return document.GetPages().Count();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to determine PDF page count.", ex);
            }
        }

        /// <summary>
        /// Extracts text for each page and returns an array where index 0 = page 1 text, index 1 = page 2, etc.
        /// </summary>
        public static string[] ExtractTextPerPage(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("PDF file not found.", filePath);

            try
            {
                using var document = PdfDocument.Open(filePath);
                var pages = document.GetPages().OrderBy(p => p.Number).ToArray();
                var result = new string[pages.Length];
                for (int i = 0; i < pages.Length; i++)
                {
                    result[i] = pages[i].Text ?? string.Empty;
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract per-page text from PDF file.", ex);
            }
        }

        /// <summary>
        /// Extracts text from a specific page (1-based pageNumber).
        /// </summary>
        /// <param name="filePath">PDF file path.</param>
        /// <param name="pageNumber">1-based page index.</param>
        /// <returns>Text for the requested page; empty string when page has no text.</returns>
        public static string ExtractTextFromPage(string filePath, int pageNumber)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("PDF file not found.", filePath);

            try
            {
                using var document = PdfDocument.Open(filePath);
                var page = document.GetPage(pageNumber);
                return page?.Text ?? string.Empty;
            }
            catch (PdfDocumentFormatException ex)
            {
                throw new InvalidOperationException("PDF format error while extracting page text.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract page text from PDF file.", ex);
            }
        }

        /// <summary>
        /// Async wrapper to extract all text from a file.
        /// </summary>
        public static Task<string> ExtractTextFromFileAsync(string filePath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ExtractTextFromFile(filePath);
            }, cancellationToken);
        }

        /// <summary>
        /// Async wrapper to extract per-page text array.
        /// </summary>
        public static Task<string[]> ExtractTextPerPageAsync(string filePath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ExtractTextPerPage(filePath);
            }, cancellationToken);
        }

        /// <summary>
        /// Converts plain extracted text to a minimal HTML document suitable for rendering inside a WebBrowser/WebView.
        /// </summary>
        public static string ConvertTextToSimpleHtml(string text, string title = "PDF Text")
        {
            if (text == null) return string.Empty;

            static string HtmlEncode(string s)
            {
                return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
            }

            var escaped = HtmlEncode(text);
            // Preserve line breaks using <pre> for simplicity.
            return $@"<!doctype html>
                    <html>
                    <head>
                      <meta charset=""utf-8"" />
                      <title>{HtmlEncode(title)}</title>
                      <style>body {{ font-family: Segoe UI, Arial, sans-serif; margin:12px; white-space:pre-wrap; }}</style>
                    </head>
                    <body>
                    <pre>{escaped}</pre>
                    </body>
                    </html>";
        }

        /// <summary>
        /// Converts a single page's text to a small HTML document, adding a page header.
        /// </summary>
        public static string ConvertPageToHtml(string pageText, int pageNumber, string title = "PDF Page")
        {
            var header = $"Page {pageNumber}";
            var body = string.IsNullOrEmpty(pageText) ? "(no text on page)" : pageText;
            var docTitle = $"{title} - {header}";
            return ConvertTextToSimpleHtml($"{header}\n\n{body}", docTitle);
        }

        //
        // New: URL-based PDF support
        //

        /// <summary>
        /// Downloads a PDF from the provided HTTP/HTTPS URL and extracts all text (async).
        /// </summary>
        /// <param name="url">An absolute HTTP/HTTPS URL pointing to a PDF.</param>
        public static async Task<string> ExtractTextFromUrlAsync(string url, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));
            }

            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var bytes = await DownloadPdfBytesFromUrlAsync(s_httpClient, uri, cancellationToken).ConfigureAwait(false);
                return ExtractTextFromBytes(bytes);
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("Failed to download PDF from URL.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract text from PDF at URL.", ex);
            }
        }

        /// <summary>
        /// Downloads a PDF from the provided HTTP/HTTPS URL and returns per-page text array (async).
        /// </summary>
        public static async Task<string[]> ExtractTextPerPageFromUrlAsync(string url, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));
            }

            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var bytes = await DownloadPdfBytesFromUrlAsync(s_httpClient, uri, cancellationToken).ConfigureAwait(false);

                // Reuse same logic as ExtractTextPerPage by opening the bytes as a stream.
                using var ms = new MemoryStream(bytes);
                using var document = PdfDocument.Open(ms);
                var pages = document.GetPages().OrderBy(p => p.Number).ToArray();
                var result = new string[pages.Length];
                for (int i = 0; i < pages.Length; i++)
                {
                    result[i] = pages[i].Text ?? string.Empty;
                }

                return result;
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("Failed to download PDF from URL.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract per-page text from PDF at URL.", ex);
            }
        }

        /// <summary>
        /// Downloads a PDF from URL and returns the number of pages (async).
        /// </summary>
        public static async Task<int> GetPageCountFromUrlAsync(string url, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));
            }

            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var bytes = await DownloadPdfBytesFromUrlAsync(s_httpClient, uri, cancellationToken).ConfigureAwait(false);

                using var ms = new MemoryStream(bytes);
                using var document = PdfDocument.Open(ms);
                return document.GetPages().Count();
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("Failed to download PDF from URL.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to determine PDF page count at URL.", ex);
            }
        }

        /// <summary>
        /// Downloads a PDF from URL and extracts text for a specific 1-based page number (async).
        /// </summary>
        public static async Task<string> ExtractTextFromUrlPageAsync(string url, int pageNumber, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber));

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));
            }

            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var bytes = await DownloadPdfBytesFromUrlAsync(s_httpClient, uri, cancellationToken).ConfigureAwait(false);

                using var ms = new MemoryStream(bytes);
                using var document = PdfDocument.Open(ms);
                var page = document.GetPage(pageNumber);
                return page?.Text ?? string.Empty;
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("Failed to download PDF from URL.", ex);
            }
            catch (PdfDocumentFormatException ex)
            {
                throw new InvalidOperationException("PDF format error while extracting page text.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract page text from PDF at URL.", ex);
            }
        }

        /// <summary>
        /// Synchronous convenience wrapper to extract all text from a PDF at URL.
        /// </summary>
        public static string ExtractTextFromUrl(string url, CancellationToken cancellationToken = default)
        {
            return ExtractTextFromUrlAsync(url, cancellationToken).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Synchronous convenience wrapper to extract per-page text array from a PDF at URL.
        /// </summary>
        public static string[] ExtractTextPerPageFromUrl(string url, CancellationToken cancellationToken = default)
        {
            return ExtractTextPerPageFromUrlAsync(url, cancellationToken).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Downloads a PDF directly or resolves a page containing an embedded PDF.
        /// Supports iframe/embed/object URLs and relative PDF links.
        /// </summary>
        public static async Task<byte[]> DownloadPdfBytesFromUrlAsync(string url, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));

            return await DownloadPdfBytesFromUrlAsync(s_httpClient, uri, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Downloads a PDF directly or through an embedded viewer using the caller's HttpClient.
        /// </summary>
        public static async Task<byte[]> DownloadPdfBytesFromUrlAsync(
            HttpClient httpClient,
            string url,
            CancellationToken cancellationToken = default)
        {
            if (httpClient == null) throw new ArgumentNullException(nameof(httpClient));
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));

            return await DownloadPdfBytesFromUrlAsync(httpClient, uri, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<byte[]> DownloadPdfBytesFromUrlAsync(
            HttpClient httpClient,
            Uri uri,
            CancellationToken cancellationToken,
            int depth = 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var response = await httpClient.GetAsync(uri, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
            var bytes = await response.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);

            if (IsPdfBytes(bytes))
                return bytes;

            if (depth >= MaxEmbeddedPdfDepth)
                throw new InvalidOperationException("The URL did not resolve to a PDF after following embedded PDF links.");

            var mediaType = response.Content.Headers.ContentType?.MediaType;
            var isHtml = string.IsNullOrWhiteSpace(mediaType) || mediaType.Contains("html", StringComparison.OrdinalIgnoreCase) || LooksLikeHtml(bytes);
            if (!isHtml)
                throw new InvalidOperationException($"The URL did not return a PDF. Content-Type: {mediaType ?? "(none)"}");

            var html = Encoding.UTF8.GetString(bytes);
            foreach (var embeddedUrl in FindEmbeddedPdfUrls(html, uri))
            {
                try
                {
                    return await DownloadPdfBytesFromUrlAsync(httpClient, embeddedUrl, cancellationToken, depth + 1).ConfigureAwait(false);
                }
                catch (HttpRequestException)
                {
                    // A page can contain several viewers; try the next candidate.
                }
                catch (InvalidOperationException)
                {
                    // Candidate may be another viewer page rather than the PDF itself.
                }
            }

            throw new InvalidOperationException("No embedded PDF URL was found in the page.");
        }

        private static bool IsPdfBytes(byte[] bytes)
        {
            return bytes.Length >= 5 && bytes[0] == '%' && bytes[1] == 'P' && bytes[2] == 'D' && bytes[3] == 'F' && bytes[4] == '-';
        }

        private static bool LooksLikeHtml(byte[] bytes)
        {
            var preview = Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 512));
            return preview.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   preview.IndexOf("<iframe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   preview.IndexOf("<embed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   preview.IndexOf("<object", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static IEnumerable<Uri> FindEmbeddedPdfUrls(string html, Uri pageUri)
        {
            var candidates = new List<string>();
            const string quotedValue = "(?<value>[^\"']+)";
            var attributePattern = $"<(?:iframe|embed|object|param)\\b[^>]*?(?:src|data|value)\\s*=\\s*[\"']{quotedValue}[\"']";
            foreach (Match match in Regex.Matches(html, attributePattern, RegexOptions.IgnoreCase))
                candidates.Add(WebUtility.HtmlDecode(match.Groups["value"].Value));

            // A direct PDF link is also commonly rendered as an anchor on viewer pages.
            foreach (Match match in Regex.Matches(html, $"<a\\b[^>]*?href\\s*=\\s*[\"']{quotedValue}[\"']", RegexOptions.IgnoreCase))
            {
                var value = WebUtility.HtmlDecode(match.Groups["value"].Value);
                if (value.Contains(".pdf", StringComparison.OrdinalIgnoreCase) ||
                    value.Contains("pdf", StringComparison.OrdinalIgnoreCase))
                    candidates.Add(value);
            }

            // Some viewers put the PDF in a JavaScript/config value instead of an HTML attribute.
            foreach (Match match in Regex.Matches(html, "(?<value>(?:https?:)?//[^\"'\\s<>]+\\.pdf(?:\\?[^\"'\\s<>]*)?|[^\"'\\s<>]+\\.pdf(?:\\?[^\"'\\s<>]*)?)", RegexOptions.IgnoreCase))
                candidates.Add(WebUtility.HtmlDecode(match.Groups["value"].Value));

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var candidateCount = 0;
            foreach (var candidate in candidates)
            {
                if (candidateCount++ >= MaxEmbeddedPdfCandidates)
                    yield break;

                var value = Uri.UnescapeDataString(candidate.Trim());
                if (value.StartsWith("//", StringComparison.Ordinal))
                    value = pageUri.Scheme + ":" + value;
                if (Uri.TryCreate(pageUri, value, out var resolved) &&
                    (resolved.Scheme == Uri.UriSchemeHttp || resolved.Scheme == Uri.UriSchemeHttps) &&
                    seen.Add(resolved.AbsoluteUri))
                    yield return resolved;
            }
        }

        /// <summary>
        /// Extracts per-page text using the caller's HttpClient so its timeout applies to embedded pages.
        /// </summary>
        public static async Task<string[]> ExtractTextPerPageFromUrlAsync(
            string url,
            HttpClient httpClient,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            if (httpClient == null) throw new ArgumentNullException(nameof(httpClient));
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
                throw new ArgumentException("URL must be an absolute HTTP/HTTPS URI.", nameof(url));

            try
            {
                var bytes = await DownloadPdfBytesFromUrlAsync(httpClient, uri, cancellationToken).ConfigureAwait(false);
                using var ms = new MemoryStream(bytes);
                using var document = PdfDocument.Open(ms);
                return document.GetPages().OrderBy(p => p.Number).Select(p => p.Text ?? string.Empty).ToArray();
            }
            catch (HttpRequestException ex)
            {
                throw new InvalidOperationException("Failed to download PDF from URL.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to extract per-page text from PDF at URL.", ex);
            }
        }
    }
}
