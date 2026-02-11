using ClosedXML.Excel;
using System.Data;

namespace ExcelControls
{
    public class ExcelController
    {
        /// <summary>
        /// Creates a new Excel workbook and saves it to the specified path
        /// </summary>
        public void CreateWorkbook(string filePath, string sheetName = "Sheet1")
        {
            using var workbook = new XLWorkbook();
            workbook.Worksheets.Add(sheetName);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Writes data to an Excel file from a DataTable
        /// </summary>
        public void WriteDataTable(string filePath, DataTable dataTable, string sheetName = "Sheet1")
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);
            worksheet.Cell(1, 1).InsertTable(dataTable);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Writes data to an Excel file from a collection of objects
        /// </summary>
        public void WriteData<T>(string filePath, IEnumerable<T> data, string sheetName = "Sheet1")
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);
            worksheet.Cell(1, 1).InsertTable(data);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Reads data from an Excel file and returns as a DataTable
        /// </summary>
        public DataTable ReadAsDataTable(string filePath, string sheetName = "Sheet1", bool firstRowAsHeader = true)
        {
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(sheetName);
            var range = worksheet.RangeUsed();

            var dataTable = new DataTable();

            if (firstRowAsHeader)
            {
                var firstRow = range.FirstRow();
                foreach (var cell in firstRow.CellsUsed())
                {
                    dataTable.Columns.Add(cell.Value.ToString());
                }

                foreach (var row in range.RowsUsed().Skip(1))
                {
                    var dataRow = dataTable.NewRow();
                    var i = 0;
                    foreach (var cell in row.CellsUsed())
                    {
                        dataRow[i++] = cell.Value.ToString();
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            else
            {
                var columnCount = range.ColumnCount();
                for (int i = 0; i < columnCount; i++)
                {
                    dataTable.Columns.Add($"Column{i + 1}");
                }

                foreach (var row in range.RowsUsed())
                {
                    var dataRow = dataTable.NewRow();
                    var i = 0;
                    foreach (var cell in row.CellsUsed())
                    {
                        dataRow[i++] = cell.Value.ToString();
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Appends data to an existing Excel file
        /// </summary>
        public void AppendData<T>(string filePath, IEnumerable<T> data, string sheetName = "Sheet1")
        {
            using var workbook = File.Exists(filePath)
                ? new XLWorkbook(filePath)
                : new XLWorkbook();

            IXLWorksheet worksheet;
            if (workbook.Worksheets.Contains(sheetName))
            {
                worksheet = workbook.Worksheet(sheetName);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                worksheet.Cell(lastRow + 1, 1).InsertData(data);
            }
            else
            {
                worksheet = workbook.Worksheets.Add(sheetName);
                worksheet.Cell(1, 1).InsertTable(data);
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Formats a worksheet with styling
        /// </summary>
        public void FormatWorksheet(string filePath, string sheetName, Action<IXLWorksheet> formatAction)
        {
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(sheetName);
            formatAction(worksheet);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Creates a styled report with header formatting
        /// </summary>
        public void CreateStyledReport<T>(string filePath, IEnumerable<T> data, string sheetName = "Report", string title = "Report")
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Add title
            worksheet.Cell(1, 1).Value = title;
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 1).Style.Font.FontSize = 16;

            // Insert data table
            var table = worksheet.Cell(3, 1).InsertTable(data);
            table.Theme = XLTableTheme.TableStyleMedium2;

            // Auto-fit columns
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Exports multiple sheets to a single workbook
        /// </summary>
        public void CreateMultiSheetWorkbook(string filePath, Dictionary<string, DataTable> sheets)
        {
            using var workbook = new XLWorkbook();

            foreach (var sheet in sheets)
            {
                var worksheet = workbook.Worksheets.Add(sheet.Key);
                worksheet.Cell(1, 1).InsertTable(sheet.Value);
                worksheet.Columns().AdjustToContents();
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Exports custom data with column definitions and color formatting
        /// </summary>
        public void ExportWithCustomColumns<T>(
            string filePath,
            Dictionary<string, List<T>> tabData,
            ColumnDefinition[] columnDefinitions,
            Func<T, string, object?> propertyGetter,
            Func<T, string, XLColor?> colorProvider,
            Action<string>? logger = null)
        {
            using var workbook = new XLWorkbook();

            foreach (var (tabName, items) in tabData)
            {
                if (items == null || items.Count == 0)
                {
                    logger?.Invoke($"Skipping empty tab: {tabName}");
                    continue;
                }

                var safeSheetName = MakeSafeSheetName(tabName);
                logger?.Invoke($"Exporting tab: {tabName} as worksheet: {safeSheetName}");

                var worksheet = workbook.Worksheets.Add(safeSheetName);

                // Write headers
                for (int col = 0; col < columnDefinitions.Length; col++)
                {
                    var cell = worksheet.Cell(1, col + 1);
                    cell.Value = columnDefinitions[col].DisplayName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Write data rows
                int row = 2;
                foreach (var item in items)
                {
                    for (int col = 0; col < columnDefinitions.Length; col++)
                    {
                        var cell = worksheet.Cell(row, col + 1);
                        var colDef = columnDefinitions[col];

                        try
                        {
                            // Get the value (this will be the link URL for hyperlink columns)
                            var value = propertyGetter(item, colDef.PropertyName);

                            if (value is bool boolValue)
                            {
                                cell.Value = boolValue ? "Yes" : "No";
                            }
                            else if (value != null)
                            {
                                var stringValue = value.ToString();

                                // Format hyperlink column
                                if (colDef.IsHyperlink)
                                {
                                    // Get the display text (e.g., "ER-123" for Code Report No)
                                    var displayText = colDef.DisplayText?.Invoke(item);

                                    if (!string.IsNullOrWhiteSpace(displayText))
                                    {
                                        // Set the cell value to the display text (what user sees)
                                        cell.Value = displayText;

                                        // Only set hyperlink if we have a valid URL
                                        if (!string.IsNullOrWhiteSpace(stringValue) &&
                                            Uri.IsWellFormedUriString(stringValue, UriKind.Absolute))
                                        {
                                            cell.SetHyperlink(new XLHyperlink(stringValue));
                                            cell.Style.Font.FontColor = XLColor.Blue;
                                            cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                                        }
                                    }
                                    else
                                    {
                                        // Fallback to the value itself if no display text
                                        cell.Value = stringValue;
                                    }
                                }
                                else
                                {
                                    cell.Value = stringValue;
                                }
                            }
                            else
                            {
                                cell.Value = string.Empty;
                            }

                            if (colDef.HasColor)
                            {
                                var color = colorProvider(item, colDef.PropertyName);
                                if (color != null)
                                {
                                    cell.Style.Fill.BackgroundColor = color;
                                }
                                // Note: If color is null, Excel will use default (white) background
                            }

                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                            // Apply alignment
                            if (colDef.CenterAlign)
                            {
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger?.Invoke($"Warning: Failed to get property {colDef.PropertyName} for row {row}: {ex.Message}");
                            cell.Value = string.Empty;
                        }
                    }
                    row++;
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                // Freeze header row
                worksheet.SheetView.FreezeRows(1);

                logger?.Invoke($"Exported {items.Count} rows to worksheet: {safeSheetName}");
            }

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Gets all sheet names from an Excel file
        /// </summary>
        public List<string> GetSheetNames(string filePath)
        {
            using var workbook = new XLWorkbook(filePath);
            return workbook.Worksheets.Select(ws => ws.Name).ToList();
        }

        /// <summary>
        /// Deletes a worksheet from an Excel file
        /// </summary>
        public void DeleteWorksheet(string filePath, string sheetName)
        {
            using var workbook = new XLWorkbook(filePath);
            if (workbook.Worksheets.Contains(sheetName))
            {
                workbook.Worksheet(sheetName).Delete();
                workbook.SaveAs(filePath);
            }
        }

        /// <summary>
        /// Copies a worksheet within the same workbook
        /// </summary>
        public void CopyWorksheet(string filePath, string sourceSheetName, string newSheetName)
        {
            using var workbook = new XLWorkbook(filePath);
            var sourceSheet = workbook.Worksheet(sourceSheetName);
            sourceSheet.CopyTo(newSheetName);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Creates a color based on value comparison logic
        /// </summary>
        public static XLColor? GetColorFromComparison(string? oldValue, string? newValue)
        {
            if (string.IsNullOrWhiteSpace(oldValue) && string.IsNullOrWhiteSpace(newValue))
            {
                return XLColor.White;
            }

            if (string.IsNullOrWhiteSpace(oldValue))
            {
                return XLColor.White;
            }

            if (string.IsNullOrWhiteSpace(newValue))
            {
                return XLColor.White;
            }

            if (string.Equals(oldValue.Trim(), newValue.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return XLColor.White;
            }

            return XLColor.Yellow;
        }

        /// <summary>
        /// Makes a string safe for use as an Excel sheet name
        /// </summary>
        public static string MakeSafeSheetName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "Sheet";

            // Excel sheet name restrictions: max 31 chars, no: / \ ? * [ ]
            var invalid = new[] { '/', '\\', '?', '*', '[', ']', ':' };
            var safe = new string(input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());

            // Trim to 31 characters
            if (safe.Length > 31)
                safe = safe.Substring(0, 31);

            safe = safe.Trim();

            return string.IsNullOrEmpty(safe) ? "Sheet" : safe;
        }
    }

    /// <summary>
    /// Defines a column for Excel export
    /// </summary>
    public class ColumnDefinition
    {
        public string PropertyName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public bool HasColor { get; set; }
        public bool IsHyperlink { get; set; }
        public bool CenterAlign { get; set; }
        public Func<object, string>? DisplayText { get; set; }
    }
}