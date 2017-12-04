namespace RobsonRocha.Exemplos.OpenXml
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;

    internal static class Utility
    {
        internal static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue?.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
            }
            return value ?? string.Empty;
        }

        internal static void SetCellValue(SpreadsheetDocument doc, Cell cell, object value)
        {
            string newValue = value.ToString();
            switch (value)
            {
                case sbyte sb:
                case byte b:
                case char c:
                case short s:
                case ushort us:
                case int i:
                case uint ui:
                case long l:
                case ulong ul:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;

                case float f:
                    newValue = ((float)value).ToString(CultureInfo.InvariantCulture);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;

                case double d:
                    newValue = ((double)value).ToString(CultureInfo.InvariantCulture);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;

                case decimal m:
                    newValue = ((decimal)value).ToString(CultureInfo.InvariantCulture);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;

                case bool bl:
                    newValue = ((bool)value) ? "1" : "0";
                    cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                    break;

                case DateTime dt:
                    newValue = ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;

                default:
                    doc.WorkbookPart.SharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(newValue)));
                    int newIndex = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.Count - 1;
                    newValue = newIndex.ToString();
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
            }
            cell.CellValue = new CellValue(newValue);
        }

        private static void ReplaceRowIndex(Row row, string curRowIndex, string newRowIndex)
        {
            foreach (Cell cell in row.Elements<Cell>())
                cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex));
        }

        private static void MoveRowsDown(Worksheet worksheet, uint rowIndex)
        {
            //TODO: Recalcular referências para as células movidas
            //TODO: Recalcular cadeia de racálculo para as células movidas
            //TODO: Recalcular hyperlinks para as células movidas
            //TODO: Recalcular ranges que interseccionem as células movidas
            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            foreach (Row row in rows)
            {
                uint newIndex = row.RowIndex + 1;
                ReplaceRowIndex(row, row.RowIndex.ToString(), newIndex.ToString());
                row.RowIndex = newIndex;
            }
        }

        internal static void InsertRow(Worksheet worksheet, Row refRow, Row insertedRow)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            uint rowIndex = refRow.RowIndex;

            MoveRowsDown(worksheet, rowIndex);

            Row newRow = sheetData.InsertBefore(insertedRow, refRow);

            string newIndex = rowIndex.ToString();

            ReplaceRowIndex(newRow, newRow.RowIndex, newIndex.ToString());

            newRow.RowIndex = rowIndex;
        }
    }
}
