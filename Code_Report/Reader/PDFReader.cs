using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace Code_Report
{
    public class PDFReader
    {
        private string _filePath;
        private bool _isRead;
        PdfReader _reader;

        public PDFReader(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("Path is not Invalid");
            }
            _filePath = filePath;
            _isRead = false;
            _reader = new PdfReader(filePath);
        }
        ~PDFReader()
        {
            if (_isRead)
            {
                _reader.Close();
            }
        }
        /// <summary>
        /// Get text page from whole document.
        /// </summary>
        public string getPDFText()
        {
            string allText = "";
            for (int page = 1; page <= this._reader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                string textPage = PdfTextExtractor.GetTextFromPage(this._reader, page, strategy);
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
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            textPage = PdfTextExtractor.GetTextFromPage(this._reader, page, strategy);
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
                for (int page = 1; page <= this._reader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string textPage = PdfTextExtractor.GetTextFromPage(this._reader, page, strategy);
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
        public List<string> getTextFrom2String(string beginText, string endText, string pages)
        {
            List<string> textCollect = new List<string>();
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
            MatchCollection matches = regex.Matches(sourceText);
            foreach (Match match in matches)
            {
                textCollect.Append(match.Groups[1].Value);
            }
            return textCollect;
        }
    }
}
