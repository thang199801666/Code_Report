using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Vml;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Reader
{
    public class ExcelReader
    {
        private string _xlPath;
        System.IO.FileStream _stream;
        XLWorkbook _workbook;
        IXLWorksheet _worksheet;

        public string XlPath { get => _xlPath; set => _xlPath = value; }

        public ExcelReader(string filePath, System.IO.FileAccess accessible = System.IO.FileAccess.Read)
        {
            _xlPath = filePath;
            if (System.IO.File.Exists(_xlPath))
            {
                _stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open, accessible, System.IO.FileShare.ReadWrite);
                _workbook = new XLWorkbook(_stream);
                _worksheet = _workbook.Worksheet(1);
            }
        }

        public ExcelReader()
        {
            _workbook = new XLWorkbook();
        }

        ~ExcelReader()
        {
            if (_stream != null)
            {
                _stream.Close();
            }
            if (_worksheet != null)
            {
                _worksheet = null;
            }
            if (_workbook != null)
            {
                _workbook.Dispose();
                _workbook = null;
            }
        }

        private static string ColumnIntToString(int columnNumber)
        {
            string columnName = columnNumber > 26 ? Convert.ToChar(64 + (columnNumber / 26)).ToString() + Convert.ToChar(64 + (columnNumber % 26)) : Convert.ToChar(64 + columnNumber).ToString();
            return columnName.ToUpper();
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

        public void SelectSheet(string name)
        {
            _worksheet = _workbook.Worksheet(name);
        }

        public void SaveWorkBook(bool saveOptions = true)
        {
            _workbook.Save(saveOptions);
        }

        public void SaveAs(string path)
        {
            _xlPath = path;
            _workbook.SaveAs(path);
        }

        public void ShowExcel()
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "Excel.exe";
            myProcess.StartInfo.Arguments = "\"" + _xlPath + "\"";
            myProcess.Start();
        }

        public int MaxRow()
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

        public void WriteCellData(int row, int col, string value)
        {
            _worksheet.Cell(row, col).Value = value;
        }

        public void WriteCellColor(int row, int col, System.Drawing.Color color)
        {
            _worksheet.Cell(row, col).Style.Fill.SetBackgroundColor(XLColor.FromColor(color));
        }

        public void SetCellColor(int row, int col, System.Drawing.Color color)
        {
            _worksheet.Cell(row, col).Style.Fill.SetBackgroundColor(XLColor.FromColor(color));
        }

        public void WriteHyperLink(int row, int col, string link)
        {
            _worksheet.Cell(row, col).SetHyperlink(new XLHyperlink(link));
        }

        public void AddWorkSheet(string sheetName)
        {
            if (!CheckContainSheet(sheetName))
            _worksheet = _workbook.Worksheets.Add(sheetName);
        }

        public bool CheckContainSheet(string sheetName)
        {
            return _workbook.Worksheets.Contains(sheetName) ? true: false;
        }

        public string GetCellIDByText(string sheetName, string text)
        {
            if (_workbook.Worksheets.Contains(sheetName))
            {
                _worksheet = _workbook.Worksheet(sheetName);
                IXLCells cell = _worksheet.Search(text, CompareOptions.OrdinalIgnoreCase);
                int minRowID = cell.First().Address.RowNumber;
                int maxRowID = _worksheet.RangeUsed().RowCount();
                int minColID = cell.First().Address.ColumnNumber;
                int maxColID = minColID + 6;
                return ColumnIntToString(minRowID) + minColID.ToString() + ":" + ColumnIntToString(maxColID) + maxColID.ToString();
            }
            else 
            {
                return "Not Found";
            }
        }

        public List<string> GetSheetNames()
        {
            List<string> names = new List<string>();
            foreach (IXLWorksheet worksheet in _workbook.Worksheets)
            {
                names.Add(worksheet.Name); 
            }
            return names;
        }
    }
}
