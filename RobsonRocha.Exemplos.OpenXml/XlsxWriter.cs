namespace RobsonRocha.Exemplos.OpenXml
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    public class WriteXlsxOptions
    {
        public string SheetName { get; set; }

        public bool Ignore { get; set; }

        public IDictionary<string, object> IndividualReplacements { get; set; }

        public IEnumerable<string> LineLocators { get; set; }

        public IEnumerable<IDictionary<string, object>> LineReplacements { get; set; }
    }

    public class XlsxWriter
    {
        public void WriteXlsx(string xlsxTemplatePath, string xlsxOutputPath, WriteXlsxOptions[] options)
        {
            File.Copy(xlsxTemplatePath, xlsxOutputPath);
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsxOutputPath, true))
            {
                foreach (Sheet sheet in doc.WorkbookPart.Workbook.Sheets.OfType<Sheet>())
                {
                    SheetInfo sheetInfo = new SheetInfo(sheet.Name);

                    foreach(WriteXlsxOptions option in options.Where(o => o.SheetName == sheetInfo.Name && 
                                                                          !o.Ignore &&
                                                                          (
                                                                            (o.IndividualReplacements?.Any() ?? false) ||
                                                                            (
                                                                                (o.LineLocators?.Any() ?? false) && 
                                                                                (o.LineReplacements?.Any() ?? false)
                                                                            )
                                                                          )))
                    {
                        Worksheet worksheet = ((WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id.Value)).Worksheet;
                        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

                        bool linesReplaced = false;

                        foreach (Row row in rows)
                        {
                            foreach (Cell cell in row.Descendants<Cell>())
                            {
                                string cellValue = Utility.GetCellValue(doc, cell);
                                if (string.IsNullOrWhiteSpace(cellValue))
                                    continue;

                                if (option.IndividualReplacements != null && option.IndividualReplacements.TryGetValue(cellValue, out object replacementValue))
                                {
                                    Utility.SetCellValue(doc, cell, replacementValue);
                                }

                                if(!linesReplaced && option.LineLocators != null && option.LineLocators.Contains(cellValue))
                                {
                                    linesReplaced = true;
                                    foreach (IDictionary<string, object> lineReplacements in option.LineReplacements)
                                    {
                                        Row cloneRow = (Row)row.CloneNode(true);
                                        foreach (var cloneCell in cloneRow.Descendants<Cell>())
                                        {
                                            string cloneCellValue = Utility.GetCellValue(doc, cloneCell);
                                            if (string.IsNullOrWhiteSpace(cloneCellValue))
                                                continue;

                                            if (lineReplacements.TryGetValue(cloneCellValue, out object cloneCellReplacementValue))
                                            {
                                                Utility.SetCellValue(doc, cloneCell, cloneCellReplacementValue);
                                            }
                                        }
                                        Utility.InsertRow(worksheet, row, cloneRow);
                                    }
                                    row.Remove();
                                }
                            }
                        }
                    }
                }
                doc.Save();
            }
        }
    }
}
