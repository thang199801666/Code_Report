using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
        private Workbook _workBook;

        private Microsoft.Office.Interop.Excel.Application _xlApp;
        public ExcelReader()
        {
            _xlApp = new Microsoft.Office.Interop.Excel.Application();
            _xlApp.Visible = false;
            _xlApp.ScreenUpdating = false;
            _xlApp.DisplayAlerts = false;
        }
        ~ExcelReader()
        {
            try
            {
                _workBook.Close(false);
                Marshal.ReleaseComObject(_workBook);
            }
            catch { }
            try
            {
                _xlApp.Visible = true;
                _xlApp.ScreenUpdating = true;
                _xlApp.DisplayAlerts = true;
                _xlApp.Quit();
                Marshal.ReleaseComObject(_xlApp);
            }
            catch { }
            GC.Collect();
        }

        public void loadWorkBook(string xlFile)
        {
            try
            {
                _xlPath = xlFile;
                _workBook = _xlApp.Workbooks.Open(xlFile, true);
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
            }
        }
        public Dictionary<string, Codes> getTableByRange(string sheetName, string range)
        {
            Dictionary<string, Codes> codeReports = new Dictionary<string, Codes>();
            List<string> types = new List<string> { "ER", "ESR", "RR" };
            List<string> webTypes = new List<string> { "IAPMO UES ER", "ICC-ES ESR", "LADBS RR" };
            Worksheet sheet = _workBook.Sheets[sheetName];//.Worksheets[sheetName];
            Range targetRange = sheet.Range[range];
            string webType = "";
            foreach (Range row in targetRange.Rows)
            {
                string ID = targetRange.Cells[row.Row, 1].Text;
                if (targetRange.Cells[row.Row, 0].Text == "")
                {
                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        string val = targetRange.Cells[row.Row, i + 1].Text.Trim();
                        webTypes.Any(x => val.Contains(x));
                        if (webTypes.Contains(val) && val != null)
                        {
                            webType = val.Trim();
                            break;
                        }
                    }
                }
                if (types.Any(w => ID.Contains(w))) //(Enum.IsDefined(typeof(Types), ID))
                {
                    Codes code = new Codes();
                    code.Number = ID.Trim();
                    foreach (var link in targetRange.Cells[row.Row, 1].Hyperlinks)
                    {
                        code.Link = link.Address;
                    }
                    code.ProductCategory = targetRange.Cells[row.Row, 2].Text.Trim();
                    code.Description = targetRange.Cells[row.Row, 3].Text.Trim();
                    code.ProductsListed = targetRange.Cells[row.Row, 4].Text.Trim();
                    code.LatestCode = targetRange.Cells[row.Row, 5].Text.Trim();
                    code.IssueDate = targetRange.Cells[row.Row, 6].Text.Trim();
                    code.ExpirationDate = targetRange.Cells[row.Row, 7].Text.Trim();
                    code.WebType = webType;
                    codeReports[ID.Trim()+"-"+ webType] = code;
                }
                Marshal.ReleaseComObject(row);
            }
            return codeReports;
        }

        public void setTableByRange(string binPath)
        {
            Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
            foreach (var code in codeReports)
            {
                Console.WriteLine(code.Value.Link);
            }
        }
    }
}
