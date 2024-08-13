using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using iText.Kernel.XMP.Impl;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using DocumentFormat.OpenXml.Drawing.Charts;
using Org.BouncyCastle.Asn1.Cms;


namespace Code_Report
{
    enum Types
    {
        ER,
        ESR,
        RR,
    }
    internal class ExcelReader
    {
        private string _xlPath;
        System.IO.FileStream _stream;
        XLWorkbook _workbook;
        IXLWorksheet _worksheet;

        public ExcelReader(string filePath, System.IO.FileAccess accessible = FileAccess.Read)
        {
            _stream = new FileStream(filePath, FileMode.Open, accessible, FileShare.ReadWrite);
            _workbook = new XLWorkbook(_stream);
            _worksheet = _workbook.Worksheet(1);
        }
        ~ExcelReader()
        {
            
        }

        public void Delete()
        {
            _stream.Close();
            _worksheet = null;
            _workbook.Dispose();
            _workbook = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public void selectSheet(string name)
        {
            _worksheet = _workbook.Worksheet(name);
        }

        public void saveWorkBook(bool saveOptions = true)
        {
            _workbook.Save(saveOptions);
        }

        public int maxRow()
        {
            int maxRow = 1;
            if (_worksheet != null)
            {
                while (_worksheet.Cell(maxRow, 1).GetValue<string>() != "" && _worksheet.Cell(maxRow, 1).GetValue<string>() != null)
                {
                    maxRow++;
                }
            }
            return maxRow;
        }

        public void writeCellData(int row, int col, string value)
        {
            _worksheet.Cell(row, col).Value = value;
        }

        public void writeCellColor(int row, int col, System.Drawing.Color color)
        {
            _worksheet.Cell(row, col).Style.Fill.SetBackgroundColor(XLColor.FromColor(color) );
        }

        public void writeHyperLink(int row, int col, string link)
        {
            _worksheet.Cell(row, col).SetHyperlink(new XLHyperlink(link));
        }

        public Dictionary<string, Codes> getTableByRange(string sheetName, string range)
        {
            Dictionary<string, Codes> codeReports = new Dictionary<string, Codes>();
            _worksheet = _workbook.Worksheet(sheetName);
            IXLCells firstCell = _worksheet.Search("Code Report No", CompareOptions.OrdinalIgnoreCase);
            IXLAddress firstCellAdd = firstCell.First().Address;
            int minRow = firstCell.First().Address.RowNumber + 1;
            int maxRow = _worksheet.RangeUsed().RowCount();
            int minCol = firstCell.First().Address.ColumnNumber;
            int maxCol = firstCell.First().Address.ColumnNumber + 7;

            List<string> webTypes = new List<string> { "IAPMO UES ER", "ICC-ES ESR"};
            string webType = "";

            for (int i = minRow; i < maxRow; i++)
            {
                string ID = _worksheet.Cell(i, minCol).GetValue<string>();
                if ((ID.Contains("ER")|| ID.Contains("ESR")) && (webType == "IAPMO UES ER" || webType == "ICC-ES ESR"))
                {
                    Codes code = new Codes();
                    code.Number = ID.Trim();
                    code.Link = _worksheet.Cell(i, minCol).GetHyperlink().ExternalAddress?.ToString();
                    code.ProductCategory = _worksheet.Cell(i, minCol + 1).GetValue<string>();
                    code.Description = _worksheet.Cell(i, minCol+2).GetValue<string>();
                    code.ProductsListed = _worksheet.Cell(i, minCol + 3).GetValue<string>();
                    code.LatestCode = _worksheet.Cell(i, minCol + 4).GetValue<string>();
                    try { code.IssueDate = _worksheet.Cell(i, minCol + 5).GetValue<DateTime>().ToString("MMM-yyyy"); }
                    catch { code.IssueDate = "n/a"; }
                    try { code.ExpirationDate = _worksheet.Cell(i, minCol + 6).GetValue<DateTime>().ToString("MMM-yyyy"); }
                    catch { code.ExpirationDate = "n/a"; }
                    code.WebType = webType;
                    codeReports[ID.Trim() + "-" + webType] = code;
                }
                for (int check = 1; check < maxCol; check++)
                {
                    string val = _worksheet.Cell(i, check).GetValue<string>();
                    if (webTypes.Any(val.Contains) && val != null)
                    {
                        webType = val.Trim();
                        break;
                    }
                }
            }
            return codeReports;
        }
        public Dictionary<string, Codes> getCodeDatas(string sheetName, string trackText)
        {
            Dictionary<string, Codes> codeReports = new Dictionary<string, Codes>();
            _worksheet = _workbook.Worksheet(sheetName);
            IXLCells firstCell = _worksheet.Search(trackText, CompareOptions.OrdinalIgnoreCase);
            if (firstCell.Count()>0)
            {
                IXLAddress firstCellAdd = firstCell.First().Address;
                int minRow = firstCell.First().Address.RowNumber + 1;
                int maxRow = _worksheet.RangeUsed().RowCount();
                int minCol = firstCell.First().Address.ColumnNumber;
                int maxCol = firstCell.First().Address.ColumnNumber + 7;

                List<string> webTypes = new List<string> { "IAPMO UES ER", "ICC-ES ESR" };
                string webType = "";

                for (int i = minRow; i < maxRow; i++)
                {
                    string ID = _worksheet.Cell(i, minCol).GetValue<string>();
                    if ((ID.Contains("ER") || ID.Contains("ESR")) && (webType == "IAPMO UES ER" || webType == "ICC-ES ESR"))
                    {
                        Codes code = new Codes();
                        code.Number = ID.Trim();
                        code.Link = _worksheet.Cell(i, minCol).GetHyperlink().ExternalAddress?.ToString();
                        code.ProductCategory = _worksheet.Cell(i, minCol + 1).GetValue<string>();
                        code.Description = _worksheet.Cell(i, minCol + 2).GetValue<string>();
                        code.ProductsListed = _worksheet.Cell(i, minCol + 3).GetValue<string>();
                        code.LatestCode = _worksheet.Cell(i, minCol + 4).GetValue<string>();
                        try { code.IssueDate = _worksheet.Cell(i, minCol + 5).GetValue<DateTime>().ToString("MMM-yyyy"); }
                        catch { code.IssueDate = "n/a"; }
                        try { code.ExpirationDate = _worksheet.Cell(i, minCol + 6).GetValue<DateTime>().ToString("MMM-yyyy"); }
                        catch { code.ExpirationDate = "n/a"; }
                        code.WebType = webType;
                        codeReports[ID.Trim() + "_" + webType] = code;
                    }
                    for (int check = minCol; check < maxCol; check++)
                    {
                        string val = _worksheet.Cell(i, check)?.GetValue<string>();
                        if (webTypes.Any(val.Contains) && val != null)
                        {
                            webType = val.Trim();
                            break;
                        }
                    }
                }
            }
            return codeReports;
        }
    }
}
