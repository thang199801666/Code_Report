using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
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
        XLWorkbook _workbook;
        IXLWorksheet _worksheet;

        public ExcelReader(string filePath)
        {
            _workbook = new XLWorkbook(new FileStream(filePath,
                                        FileMode.Open,
                                        FileAccess.Read,
                                        FileShare.ReadWrite));
            _worksheet = _workbook.Worksheet(1);
        }
        ~ExcelReader()
        {
            _workbook.Dispose();
            _worksheet.Delete();
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

        public void writeHyperLink(int row, int col, string link)
        {
            //_workSheet.Hyperlinks.Add(_workSheet.Cells[row, col], link, Type.Missing);
        }

        public void getTableByRange(string sheetName)
        {
            _worksheet = _workbook.Worksheet(sheetName);
            IXLCells firstCell = _worksheet.Search("Code Report No", CompareOptions.OrdinalIgnoreCase);
            IXLAddress firstCellAdd = firstCell.First().Address;

            int minRow = firstCell.First().Address.RowNumber + 1;
            int maxRow = _worksheet.RangeUsed().RowCount();
            int minCol = firstCell.First().Address.ColumnNumber;
            int maxCol = firstCell.First().Address.ColumnNumber + 7;
            for (int i = minRow; i < maxRow; i++)
            {
                for (int j = minCol; j < maxCol; j++)
                {
                    string data = _worksheet.Cell(i, j).GetValue<string>();
                    Console.Write(data);
                    Console.Write("\t");
                }
                Console.WriteLine();
            }
            GC.Collect();
        }
    }
}
