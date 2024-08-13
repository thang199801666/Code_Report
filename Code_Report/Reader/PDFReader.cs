using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Code_Report
{
    /*
     * Purpose: A Class for customize PDF reader 
     * Author: Thang Nguyen Huu
     * Created Date: 6/9/2024
     */
    public class PDFReader
    {
        private string _filePath;
        private bool _isRead = false;
        PdfDocument _reader;

        public PDFReader(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }
            _filePath = filePath;
            _isRead = true;
            _reader = new PdfDocument(new PdfReader(filePath));
        }
        ~PDFReader()
        {
            //if (_isRead)
            //{
            //    _reader.Close();
            //}
        }
        public void Close()
        {
            _reader.Close();
        }
        /// <summary>
        /// Get text page from whole document.
        /// </summary>
        public string getPDFText()
        {
            string allText = "";
            for (int page = 1; page <= this._reader.GetNumberOfPages(); page++)
            {
                iText.Kernel.Pdf.Canvas.Parser.Listener.ITextExtractionStrategy strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                string textPage = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(this._reader.GetPage(page), strategy);
                //string textPage = PdfTextExtractor.GetTextFromPage(this._reader, page, strategy);
                allText += textPage;
            }
            return allText;
        }
        /// <summary>
        /// Get text page from 1 page.
        /// </summary>
        public string getPageText(int page)
        {
            string textPage = "";
            if (page <= _reader.GetNumberOfPages())
            {
                iText.Kernel.Pdf.Canvas.Parser.Listener.ITextExtractionStrategy strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                textPage = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(this._reader.GetPage(page), strategy);
            }
            return textPage;
        }
        /// <summary>
        /// Get text page from start and end page.
        /// </summary>
        public string getPagesText(int start, int end)
        {
            string pageText = "";
            if (end < start)
            {
                throw new Exception("End page can not smaller than start page.");
            }
            else
            {
                for (int page = 1; page <= this._reader.GetNumberOfPages(); page++)
                {
                    iText.Kernel.Pdf.Canvas.Parser.Listener.ITextExtractionStrategy strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy();
                    string textPage = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(this._reader.GetPage(page), strategy);
                    pageText += textPage;
                }
            }
            return pageText;
        }
        /// <summary>
        /// pages is range of 2 pages Ex: 1-9.
        /// pages can name as "ALL".
        /// page(s) can only one page like 1.
        /// </summary>
        public string getTextFrom2String(string beginText, string endText, string pages)
        {
            string sourceText = "";
            int noPage;
            if (pages == "All")
            {
                sourceText = getPDFText();
            }
            else if (pages.Contains("-"))
            {
                sourceText = getPagesText(Int32.Parse(pages.Split('-')[0]), Int32.Parse(pages.Split('-')[1]));
            }
            else if (int.TryParse(pages, out noPage))
            {
                sourceText = getPageText(noPage);
            }
            else
            {
                throw new Exception("Invalid input!!!");
            }
            string pattern = beginText + "(.*)" + endText;
            Regex regex = new Regex(pattern);
            Match match = regex.Match(sourceText);
            return match.Groups[1].Value.Trim();
        }
        public string getTextFrom2String_s(string pages)
        {
            string sourceText = "";
            int noPage;
            if (pages == "All")
            {
                sourceText = getPDFText();
            }
            else if (pages.Contains("-"))
            {
                sourceText = getPagesText(Int32.Parse(pages.Split('-')[0]), Int32.Parse(pages.Split('-')[1]));
            }
            else if (int.TryParse(pages, out noPage))
            {
                sourceText = getPageText(noPage);
            }
            else
            {
                throw new Exception("Invalid input!!!");
            }
            string pattern = "(?s)following codes(.*?)(\\d+)";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            Match match = regex.Match(sourceText);
            return match.Groups[2].Value.Trim();
        }
        public string getTextFromPattern(string pages, string pattern, bool ignoreCase=true)
        {
            string sourceText = "";
            int noPage;
            if (pages == "All")
            {
                sourceText = getPDFText();
            }
            else if (pages.Contains("-"))
            {
                sourceText = getPagesText(Int32.Parse(pages.Split('-')[0]), Int32.Parse(pages.Split('-')[1]));
            }
            else if (int.TryParse(pages, out noPage))
            {
                sourceText = getPageText(noPage);
            }
            else
            {
                throw new Exception("Invalid input!!!");
            }
            RegexOptions opts = new RegexOptions();
            if (ignoreCase)
            { 
                opts = RegexOptions.IgnoreCase;
            }
            Regex regex = new Regex(pattern, options:opts);
            Match match = regex.Match(sourceText);
            string output = match.Groups[1].Value.Trim();
            return output;
        }
        public Match getMatchFromPattern(string pages, string pattern, bool ignoreCase = true)
        {
            string sourceText = "";
            int noPage;
            if (pages == "All")
            {
                sourceText = getPDFText();
            }
            else if (pages.Contains("-"))
            {
                sourceText = getPagesText(Int32.Parse(pages.Split('-')[0]), Int32.Parse(pages.Split('-')[1]));
            }
            else if (int.TryParse(pages, out noPage))
            {
                sourceText = getPageText(noPage);
            }
            else
            {
                throw new Exception("Invalid input!!!");
            }
            RegexOptions opts = new RegexOptions();
            if (ignoreCase)
            {
                opts = RegexOptions.IgnoreCase;
            }
            Regex regex = new Regex(pattern, options: opts);
            Match match = regex.Match(sourceText);
            return match;
        }

        public List<string> iapmoSearch()
        {
            List<string> dateList = new List<string>();
            string latestCode = getTextFrom2String_s("1") ?? "n/a";
            string oriIssue = getTextFromPattern("1", "(?s)Originally Issued:.*?(\\d{1,2}\\/\\d{1,2}\\/\\d{2,4})");//getTextFrom2String("Originally Issued:", "Revised:", "1");
            string revised = getTextFromPattern("1", "(?s)Revised:.*?(\\d{1,2}\\/\\d{1,2}\\/\\d{2,4})");//getTextFrom2String("Revised:", "Valid Through:", "1");
            string validThrough = getTextFromPattern("1", "(?s)Valid Through:.*?(\\d{1,2}\\/\\d{1,2}\\/\\d{2,4})");//getTextFrom2String("Valid Through:", "\n", "1");

            if (revised == "" || revised == null)
            { revised = oriIssue; }

            if (latestCode != "" && latestCode != null)
            {
                dateList.Add(latestCode);
            }
            else
            {
                dateList.Add("n/a");
            }

            if (revised != "" && revised != null)
            {
                dateList.Add(IOData.convertDateTime(revised, "MM/dd/yyyy").ToString("MMM-yyyy"));
            }
            else if (revised == "" && oriIssue != "")
            { 
                revised = oriIssue;
                dateList.Add(IOData.convertDateTime(revised, "MM/dd/yyyy").ToString("MMM-yyyy"));
            }
            else
            {
                dateList.Add("n/a");
            }

            if (validThrough != "" && validThrough != null)
            {
                dateList.Add(IOData.convertDateTime(validThrough, "MM/dd/yyyy").ToString("MMM-yyyy"));
            }
            else 
            {
                dateList.Add("n/a");
            }
            return dateList;
        }

        public List<string> iccesSearch()
        {
            List<string> dateList = new List<string>();
            string latestCode = getTextFrom2String_s("1");
            Match reIssueMatch = getMatchFromPattern("1", @"(?s)issue.*?((Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(.*?)+(\d\d\d\d))");
            Match revisedMatch = getMatchFromPattern("1", @"(?s)Revised.*?((Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(.*?)+(\d\d\d\d))");
            Match validThroughMatch = getMatchFromPattern("1", @"(?s)Subject to renewal.*?((Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(.*?)(\d+))");

            string reIssued = reIssueMatch.Length > 0 ? reIssueMatch.Groups[2].Value.Trim() + " " + reIssueMatch.Groups[4].Value.Trim() : "n/a";
            string revised = revisedMatch.Length > 0 ? revisedMatch.Groups[2].Value.Trim() + " " + revisedMatch.Groups[4].Value.Trim() : "n/a";
            string validThrough = validThroughMatch.Length > 0 ? validThroughMatch.Groups[2].Value.Trim() + " " + validThroughMatch.Groups[4].Value.Trim() : "n/a";

            if (revised == "n/a" && reIssued != "n/a")
            {
                revised = reIssued;
            }
            else if (revised != "n/a" && reIssued != "n/a")
            {
                if (IOData.convertDateTime(reIssued, "MMMM yyyy") > IOData.convertDateTime(revised, "MMMM yyyy"))
                {
                    revised = reIssued;
                }
            }

            try { reIssued = IOData.convertDateTime(reIssued, "MMMMM yyyy").ToString("MMM-yyyy"); }
            catch { reIssued = "n/a"; }

            try { revised = IOData.convertDateTime(revised, "MMMMM yyyy").ToString("MMM-yyyy"); }
            catch { revised = "n/a"; }

            try { validThrough = IOData.convertDateTime(validThrough, "MMMMM yyyy").ToString("MMM-yyyy"); }
            catch { validThrough = "n/a"; }

            dateList.Add(latestCode);
            dateList.Add(revised);
            dateList.Add(validThrough);
            return dateList;
        }
    }
}
