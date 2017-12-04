namespace RobsonRocha.Exemplos.OpenXml
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Linq;
    using System.Collections.Generic;

    /// <summary>
    /// Contains methods for reading values from Xlsx files, without any formatting
    /// </summary>
    public class XlsxReader
    {
        /// <summary>
        /// Reads an Xlsx file and returns an <see cref="IReadOnlyList{SheetInfo}"/> containing the information about the sheets read
        /// </summary>
        /// <param name="xlsxPath">The path of the Xlsx file</param>
        /// <param name="options">An <see cref="IEnumerable{ReadXlsxOptions}"/> describing options for each sheet in the file</param>
        /// <returns>An <see cref="IReadOnlyList{SheetInfo}"/> containing all the data from the Xlsx file</returns>
        public IReadOnlyList<SheetInfo> ReadXlsx(string xlsxPath, IEnumerable<ReadXlsxOptions> options)
        {
            var results = new List<SheetInfo>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsxPath, false))
            {
                foreach(Sheet sheet in doc.WorkbookPart.Workbook.Sheets.OfType<Sheet>())
                {
                    SheetInfo sheetData = new SheetInfo(sheet.Name);
                    var option = options.FirstOrDefault(o => o.SheetName == sheetData.Name);
                    if (option?.Ignore ?? false)
                        continue;

                    if(option?.Headers != null)
                        sheetData.ColumnList.AddRange(option.Headers); 

                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                    foreach (Row row in rows)
                    {
                        RowInfo rowData = sheetData.AddRow();
                        int headerIndex = 0;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var cellValue = Utility.GetCellValue(doc, cell);
                            if (row.RowIndex.Value == option?.HeaderRowIndex)
                                sheetData.ColumnList.Add(cellValue);
                            string headerName = null;
                            if (row.RowIndex >= option?.StartingRowIndex)
                                headerName = sheetData.ColumnList.Count > headerIndex ? sheetData.ColumnList[headerIndex] : null;
                            rowData.AddColumn(cell.CellReference, headerName, cellValue);
                            headerIndex++;
                        }
                    }

                    results.Add(sheetData);
                }
            }

            return results;
        }
    }
}
