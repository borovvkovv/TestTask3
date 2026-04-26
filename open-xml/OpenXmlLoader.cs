using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TestTask3.interfaces;

namespace TestTask3.openXml
{
    internal class OpenXmlLoader(string path) : ITableLoader
    {
        public DataTable LoadTable(string tableName)
        {
            DataTable booksTable = new DataTable(tableName);
            using (var excelDocument = SpreadsheetDocument.Open(path, false))
            {
                var bookPart = excelDocument.WorkbookPart;
                var productsSheet = GetWorksheet(bookPart, tableName);
                var isFirstRow = true;

                foreach (var row in productsSheet.Descendants<Row>())
                {
                    if (row.Descendants<Cell>().Any(c => c.CellValue == null))
                        continue;

                    var tableRow = new List<string>();
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        tableRow.Add(GetCellValue(bookPart, productsSheet.WorksheetPart, cell));
                    }

                    if (isFirstRow)
                    {
                        foreach (var columnName in tableRow)
                            booksTable.Columns.Add(columnName);
                        isFirstRow = false;
                    }
                    else
                    {
                        DataRow newRow = booksTable.NewRow();
                        newRow.ItemArray = tableRow.ToArray();
                        booksTable.Rows.Add(newRow);
                    }
                }

                return booksTable;
            }
        }

        public bool SaveTable(DataTable dataTable, string tableName)
        {
            using (var excelDocument = SpreadsheetDocument.Open(path, true))
            {
                var bookPart = excelDocument.WorkbookPart;
                var productsSheet = GetWorksheet(bookPart, tableName);
                var rowIndex = 0;
                var isFirstRow = true;

                foreach (var row in productsSheet.Descendants<Row>())
                {
                    // Пропуск заголовка и не полностью заполненных строк.
                    if (isFirstRow || row.Descendants<Cell>().Any(c => c.CellValue == null))
                    {
                        isFirstRow = false;
                        continue;
                    }

                    var columnIndex = 0;
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        var cellValue = GetCellValue(bookPart, productsSheet.WorksheetPart, cell);
                        var datasetRow = dataTable.AsEnumerable().ElementAt(rowIndex);
                        var datasetCellValue = datasetRow.Field<string>(columnIndex);

                        if (cellValue != datasetCellValue) SetCellValue(bookPart, cell, datasetCellValue);
                        columnIndex++;
                    }

                    rowIndex++;

                    excelDocument.Save();
                }

                return true;
            }
        }

        private Worksheet GetWorksheet(WorkbookPart bookPart, string sheetName)
        {
            foreach (Sheet s in bookPart.Workbook.Descendants<Sheet>())
            {
                if (s.Name == sheetName && !string.IsNullOrEmpty(s.Id))
                {
                    if (bookPart.TryGetPartById(s.Id!, out var part))
                    {
                        if (part is WorksheetPart result)
                        {
                            return result.Worksheet;
                        }
                    }
                }
            }
            return null;
        }

        private string GetCellValue(WorkbookPart workbookPart, WorksheetPart workSheetPart, Cell cell)
        {
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }

            else
            {
                uint styleIndex = cell.StyleIndex ?? 0;
                var cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>()
                    .ElementAt((int)styleIndex);
                uint numFmtId = cellFormat.NumberFormatId.Value;
                bool isDate = numFmtId >= 14 && numFmtId <= 22;

                if (isDate)
                {
                    double number = 0;
                    var cellValueRaw = workSheetPart.Worksheet.Descendants<Cell>()
                        .Where(x => x.CellValue.InnerText == value).FirstOrDefault().InnerText;
                    if (double.TryParse(cellValueRaw, out number))
                        return DateTime.FromOADate(number).ToString("dd.MM.yyyy");
                }
            }

            return value;
        }

        private void SetCellValue(WorkbookPart workbookPart, Cell cell, string newValue)
        {
            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;
            var innerXml = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerXml;
            var oldValue = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerXml =
                innerXml.Replace(oldValue, newValue);


        }

    }
}
