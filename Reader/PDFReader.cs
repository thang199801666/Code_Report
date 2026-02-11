using AngleSharp.Dom;
using AngleSharp.Html;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace Reader
{
    /*
     * Purpose: A Class for customize PDF reader 
     * Author: Thang Nguyen Huu
     * Created Date: 6/9/2024
     *
     * CHANGES: Replaced iText-based reading with UglyToad.PdfPig to reliably extract
     * text from in-memory PDFs (downloaded from the web) and from disk.
     *
     * NOTE: Add the PdfPig NuGet package 'UglyToad.PdfPig' to the Reader project.
     */
    public class PDFReader : IDisposable
    {
        private string _filePath = string.Empty;
        private byte[] _buffer;
        private bool _isRead = false;
        bool disposed;

        public bool IsRead { get => _isRead; set => _isRead = value; }

        public PDFReader(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }
            _filePath = filePath;
            IsRead = true;
        }

        // New: construct from bytes (e.g. downloaded PDF content)
        public PDFReader(byte[] fileBytes)
        {
            if (fileBytes == null || fileBytes.Length == 0)
            {
                return;
            }

            _buffer = fileBytes;
            IsRead = true;
        }

        ~PDFReader()
        {
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //dispose managed resources
                    _buffer = null;
                }
            }
            //dispose unmanaged resources
            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Extracts the full text from the PDF using PdfPig.
        /// Supports both in-memory buffer and file path.
        /// </summary>
        public string ExtractTextFromPdf()
        {
            var text = new StringBuilder();

            if (_buffer != null)
            {
                MemoryStream ms = null;
                try
                {
                    ms = new MemoryStream(_buffer);
                    using (var pdf = PdfDocument.Open(ms))
                    {
                        foreach (var page in pdf.GetPages())
                        {
                            // PdfPig's Page.Text returns the page text
                            text.AppendLine(page.Text);
                        }
                    }
                }
                finally
                {
                    if (ms != null) ms.Close();
                }
            }
            else if (!string.IsNullOrEmpty(_filePath))
            {
                // Open directly from path
                using (var pdf = PdfDocument.Open(_filePath))
                {
                    foreach (var page in pdf.GetPages())
                    {
                        text.AppendLine(page.Text);
                    }
                }
            }

            return text.ToString();
        }

        /// <summary>
        /// Extracts only the first page's text (1-based) and returns it.
        /// Public so callers can request only first-page extraction without scanning the full document.
        /// </summary>
        public string ExtractFirstPageText()
        {
            return GetTextFromPage(1) ?? string.Empty;
        }

        /// <summary>
        /// Extracts text from a specific page (1-based) using PdfPig.
        /// </summary>
        private string GetTextFromPage(int pageNumber)
        {
            if (pageNumber <= 0) return string.Empty;

            if (_buffer != null)
            {
                MemoryStream ms = null;
                try
                {
                    ms = new MemoryStream(_buffer);
                    using (var pdf = PdfDocument.Open(ms))
                    {
                        var page = pdf.GetPage(pageNumber);
                        return page?.Text ?? string.Empty;
                    }
                }
                finally
                {
                    if (ms != null) ms.Close();
                }
            }
            else if (!string.IsNullOrEmpty(_filePath))
            {
                using (var pdf = PdfDocument.Open(_filePath))
                {
                    var page = pdf.GetPage(pageNumber);
                    return page?.Text ?? string.Empty;
                }
            }

            return string.Empty;
        }

        public string FindLatestCode()
        {
            string sourceText = GetTextFromPage(1);
            Regex latestCodeRex = new Regex("((?s)following codes.*?)(\\d+4)", RegexOptions.IgnoreCase);
            Match match = latestCodeRex.Match(sourceText);
            return match.Groups.Count > 2 ? match.Groups[2].Value.Trim() : string.Empty;
        }

        //public void iapmoCheck(ref Codes code)
        //{
        //    string sourceText = GetTextFromPage(1);
        //    Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
        //    Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
        //    Regex revisedDateRegex = new Regex(@"(?s)(revised.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
        //    Regex expirationDateRegex = new Regex(@"(renewal|valid)[\S\s\t\n]+(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        //    Match latestMatch = latestCodeRex.Match(sourceText);
        //    Match issueMatch = issueDateRegex.Match(sourceText);
        //    Match revisedMatch = revisedDateRegex.Match(sourceText);
        //    Match expirationMatch = expirationDateRegex.Match(sourceText);

        //    DateTime issueDate = new DateTime();
        //    DateTime revisedDate = new DateTime();
        //    DateTime expirationDate = new DateTime();
        //    DateTime.TryParse(issueMatch?.Groups.Count > 2 ? issueMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out issueDate);
        //    DateTime.TryParse(revisedMatch?.Groups.Count > 2 ? revisedMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out revisedDate);
        //    DateTime.TryParse(expirationMatch?.Groups.Count > 2 ? expirationMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out expirationDate);

        //    if (revisedDate > issueDate) issueDate = revisedDate;
        //    code.LatestCode = (latestMatch.Length == 0) ? "n/a" : latestMatch.Groups[1].Value;
        //    code.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
        //    code.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");
        //}

        //public void iccesCheck(ref Codes code)
        //{
        //    string sourceText = GetTextFromPage(1);
        //    Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
        //    Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
        //    Regex revisedDateRegex = new Regex(@"(?s)(Revised.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
        //    Regex expirationDateRegex = new Regex(@"(?s)(renewal|Valid.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
        //    Match latestMatch = latestCodeRex.Match(sourceText);
        //    Match issueMatch = issueDateRegex.Match(sourceText);
        //    Match revisedMatch = revisedDateRegex.Match(sourceText);
        //    Match expirationMatch = expirationDateRegex.Match(sourceText);

        //    DateTime issueDate = new DateTime();
        //    DateTime revisedDate = new DateTime();
        //    DateTime expirationDate = new DateTime();
        //    DateTime.TryParse((issueMatch != null && issueMatch.Groups.Count > 3) ? (issueMatch.Groups[2].Value.Replace(" ", "") + issueMatch.Groups[3].Value) : string.Empty, out issueDate);
        //    DateTime.TryParse((revisedMatch != null && revisedMatch.Groups.Count > 3) ? (revisedMatch.Groups[2].Value.Replace(" ", "") + revisedMatch.Groups[3].Value) : string.Empty, out revisedDate);
        //    DateTime.TryParse((expirationMatch != null && expirationMatch.Groups.Count > 3) ? (expirationMatch.Groups[2].Value.Replace(" ", "") + expirationMatch.Groups[3].Value) : string.Empty, out expirationDate);

        //    if (revisedDate > issueDate) issueDate = revisedDate;
        //    code.LatestCode = (latestMatch.Length == 0) ? "n/a" : latestMatch.Groups[1].Value;
        //    code.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
        //    code.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");
        //}


        public void Test_iapmoCheck()
        {
            string sourceText = GetTextFromPage(1);
            Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
            Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
            Regex revisedDateRegex = new Regex(@"(?s)(revised.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
            Regex expirationDateRegex = new Regex(@"(?s)(renewal|valid.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
            Match latestMatch = latestCodeRex.Match(sourceText);
            Match issueMatch = issueDateRegex.Match(sourceText);
            Match revisedMatch = revisedDateRegex.Match(sourceText);
            Match expirationMatch = expirationDateRegex.Match(sourceText);

            DateTime issueDate = new DateTime();
            DateTime revisedDate = new DateTime();
            DateTime expirationDate = new DateTime();
            DateTime.TryParse(issueMatch?.Groups.Count > 2 ? issueMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out issueDate);
            DateTime.TryParse(revisedMatch?.Groups.Count > 2 ? revisedMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out revisedDate);
            DateTime.TryParse(expirationMatch?.Groups.Count > 2 ? expirationMatch.Groups[2].Value.Replace(" ", "") : string.Empty, out expirationDate);
        }
    }
}




















//using AngleSharp.Dom;
//using AngleSharp.Html;
//using DocumentFormat.OpenXml.Spreadsheet;
//using IO;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Text;
//using System.Text.RegularExpressions;
//using iText.Kernel.Pdf;
//using iText.Kernel.Pdf.Canvas.Parser;
//using iText.Kernel.Pdf.Canvas.Parser.Listener;
//using System.Runtime.InteropServices;
//using System.Security.Cryptography;

//namespace Reader
//{
//    /*
//     * Purpose: A Class for customize PDF reader 
//     * Author: Thang Nguyen Huu
//     * Created Date: 6/9/2024
//     *
//     * CHANGES: Constructor overload accepting a byte[] so PDFs downloaded from the web
//     * can be processed directly from memory without writing a temp file.
//     *
//     * COMPAT: Updated to avoid C# 8+ features so this file builds under C# 7.3 / .NET Framework 4.7.2.
//     */
//    public class PDFReader : IDisposable
//    {
//        private string _filePath = string.Empty;
//        private byte[] _buffer;
//        private bool _isRead = false;
//        bool disposed;

//        public bool IsRead { get => _isRead; set => _isRead = value; }

//        public PDFReader(string filePath)
//        {
//            if (!File.Exists(filePath))
//            {
//                return;
//            }
//            _filePath = filePath;
//            IsRead = true;
//        }

//        // New: construct from bytes (e.g. downloaded PDF content)
//        public PDFReader(byte[] fileBytes)
//        {
//            if (fileBytes == null || fileBytes.Length == 0)
//            {
//                return;
//            }

//            _buffer = fileBytes;
//            IsRead = true;
//        }

//        ~PDFReader()
//        {
//        }
//        protected virtual void Dispose(bool disposing)
//        {
//            if (!disposed)
//            {
//                if (disposing)
//                {
//                    //dispose managed resources
//                    _buffer = null;
//                }
//            }
//            //dispose unmanaged resources
//            disposed = true;
//        }

//        public void Dispose()
//        {
//            Dispose(true);
//            GC.SuppressFinalize(this);
//        }

//        public string ExtractTextFromPdf()
//        {
//            StringBuilder text = new StringBuilder();

//            if (_buffer != null)
//            {
//                var ms = new MemoryStream(_buffer);
//                try
//                {
//                    PdfReader reader = new PdfReader(ms);
//                    try
//                    {
//                        PdfDocument pdfDoc = new PdfDocument(reader);
//                        try
//                        {
//                            SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
//                            {
//                                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);
//                                text.Append(pageText);
//                            }
//                        }
//                        finally
//                        {
//                            pdfDoc.Close();
//                        }
//                    }
//                    finally
//                    {
//                        reader.Close();
//                    }
//                }
//                finally
//                {
//                    ms.Close();
//                }
//            }
//            else
//            {
//                PdfReader reader = new PdfReader(_filePath);
//                try
//                {
//                    PdfDocument pdfDoc = new PdfDocument(reader);
//                    try
//                    {
//                        SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                        for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
//                        {
//                            string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);
//                            text.Append(pageText);
//                        }
//                    }
//                    finally
//                    {
//                        pdfDoc.Close();
//                    }
//                }
//                finally
//                {
//                    reader.Close();
//                }
//            }

//            return text.ToString();
//        }

//        private string GetTextFromPage(int pageNumber)
//        {
//            string singlePageText = string.Empty;

//            if (_buffer != null)
//            {
//                var ms = new MemoryStream(_buffer);
//                try
//                {
//                    PdfReader reader = new PdfReader(ms);
//                    try
//                    {
//                        PdfDocument pdfDoc = new PdfDocument(reader);
//                        try
//                        {
//                            SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                            singlePageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNumber), strategy);
//                        }
//                        finally
//                        {
//                            pdfDoc.Close();
//                        }
//                    }
//                    finally
//                    {
//                        reader.Close();
//                    }
//                }
//                finally
//                {
//                    ms.Close();
//                }
//            }
//            else
//            {
//                PdfReader reader = new PdfReader(_filePath);
//                try
//                {
//                    PdfDocument pdfDoc = new PdfDocument(reader);
//                    try
//                    {
//                        SimpleTextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                        singlePageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNumber), strategy);
//                    }
//                    finally
//                    {
//                        pdfDoc.Close();
//                    }
//                }
//                finally
//                {
//                    reader.Close();
//                }
//            }

//            return singlePageText;
//        }

//        public string FindLatestCode()
//        {
//            string sourceText = GetTextFromPage(1);
//            Regex latestCodeRex = new Regex("((?s)following codes.*?)(\\d+4)", RegexOptions.IgnoreCase);
//            Match match = latestCodeRex.Match(sourceText);
//            return match.Groups[2].Value.Trim();
//        }

//        public void iapmoCheck(ref Codes code)
//        {
//            string sourceText = GetTextFromPage(1);
//            Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
//            Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            Regex revisedDateRegex = new Regex(@"(?s)(revised.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            //Regex expirationDateRegex = new Regex(@"(?s)(renewal|valid.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            Regex expirationDateRegex = new Regex(@"(renewal|valid)[\S\s\t\n]+(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase | RegexOptions.Multiline);
//            Match latestMatch = latestCodeRex.Match(sourceText);
//            Match issueMatch = issueDateRegex.Match(sourceText);
//            Match revisedMatch = revisedDateRegex.Match(sourceText);
//            Match expirationMatch = expirationDateRegex.Match(sourceText);

//            DateTime issueDate = new DateTime();
//            DateTime revisedDate = new DateTime();
//            DateTime expirationDate = new DateTime();
//            DateTime.TryParse(issueMatch?.Groups[2].Value.Replace(" ", ""), out issueDate);
//            DateTime.TryParse(revisedMatch?.Groups[2].Value.Replace(" ", ""), out revisedDate);
//            DateTime.TryParse(expirationMatch?.Groups[2].Value.Replace(" ", ""), out expirationDate);

//            if (revisedDate > issueDate) issueDate = revisedDate;
//            code.LatestCode = (latestMatch.Length == 0) ? code.LatestCode = "n/a" : latestMatch?.Groups[1].Value;
//            code.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
//            code.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");
//        }

//        public void iccesCheck(ref Codes code)
//        {
//            string sourceText = GetTextFromPage(1);
//            Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
//            Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
//            Regex revisedDateRegex = new Regex(@"(?s)(Revised.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
//            Regex expirationDateRegex = new Regex(@"(?s)(renewal|Valid.*?)[\s\t\n](.*?)(\d{4})", RegexOptions.IgnoreCase);
//            Match latestMatch = latestCodeRex.Match(sourceText);
//            Match issueMatch = issueDateRegex.Match(sourceText);
//            Match revisedMatch = revisedDateRegex.Match(sourceText);
//            Match expirationMatch = expirationDateRegex.Match(sourceText);

//            DateTime issueDate = new DateTime();
//            DateTime revisedDate = new DateTime();
//            DateTime expirationDate = new DateTime();
//            DateTime.TryParse(issueMatch?.Groups[2].Value.Replace(" ", "") + issueMatch?.Groups[3].Value, out issueDate);
//            DateTime.TryParse(revisedMatch?.Groups[2].Value.Replace(" ", "") + revisedMatch?.Groups[3].Value, out revisedDate);
//            DateTime.TryParse(expirationMatch?.Groups[2].Value.Replace(" ", "") + expirationMatch?.Groups[3].Value, out expirationDate);

//            if (revisedDate > issueDate) issueDate = revisedDate;
//            code.LatestCode = (latestMatch.Length == 0) ? code.LatestCode = "n/a" : latestMatch?.Groups[1].Value;
//            code.IssueDate = (issueDate == DateTime.MinValue) ? "n/a" : issueDate.ToString("MMM-yyyy");
//            code.ExpirationDate = (expirationDate == DateTime.MinValue) ? "n/a" : expirationDate.ToString("MMM-yyyy");
//        }


//        public void Test_iapmoCheck()
//        {
//            string sourceText = GetTextFromPage(1);
//            Regex latestCodeRex = new Regex(@"(?s)following[\s\t\n]+code.*?(\d{4})", RegexOptions.IgnoreCase);
//            Regex issueDateRegex = new Regex(@"(?s)(issue.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            Regex revisedDateRegex = new Regex(@"(?s)(revised.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            Regex expirationDateRegex = new Regex(@"(?s)(renewal|valid.*?)[\s\t\n]([0-9\/]+)", RegexOptions.IgnoreCase);
//            //Regex expirationDateRegex = new Regex(@"(renewal|valid)[\S\s\t\n]+(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase | RegexOptions.Multiline);
//            Match latestMatch = latestCodeRex.Match(sourceText);
//            Match issueMatch = issueDateRegex.Match(sourceText);
//            Match revisedMatch = revisedDateRegex.Match(sourceText);
//            Match expirationMatch = expirationDateRegex.Match(sourceText);

//            DateTime issueDate = new DateTime();
//            DateTime revisedDate = new DateTime();
//            DateTime expirationDate = new DateTime();
//            DateTime.TryParse(issueMatch?.Groups[2].Value.Replace(" ", ""), out issueDate);
//            DateTime.TryParse(revisedMatch?.Groups[2].Value.Replace(" ", ""), out revisedDate);
//            DateTime.TryParse(expirationMatch?.Groups[2].Value.Replace(" ", ""), out expirationDate);
//        }
//    }
//}