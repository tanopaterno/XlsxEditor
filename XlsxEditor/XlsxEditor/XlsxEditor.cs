using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;

namespace XlsxEditor
{
    public static class XlsxEditor
    {
        public static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        public static SpreadsheetDocument GetSpreadsheetDocument(string path, bool isEditable)
        {
            var baseXls = File.ReadAllBytes(path);
            using (var ms = new MemoryStream())
            {
                ms.Write(baseXls, 0, baseXls.Length);
                return SpreadsheetDocument.Open(ms, isEditable);
            }
        }

        public static SpreadsheetDocument GetSpreadsheetDocument(MemoryStream stream, bool isEditable)
        {
            return SpreadsheetDocument.Open(stream, isEditable);
        }

        public static WorksheetPart GetWorksheetPartBySheetName(string path, string sheetName, bool isEditable)
        {
            var document = GetSpreadsheetDocument(path, isEditable);
            var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            var relationshipId = sheets.First().Id.Value;
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        public static WorksheetPart GetWorksheetPartBySheetName(this SpreadsheetDocument document, string sheetName)
        {
            var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            var relationshipId = sheets.First().Id.Value;
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        public static Worksheet GetWorksheetBySheetName(string path, string sheetName, bool isEditable)
        {
            var worksheetPart = GetWorksheetPartBySheetName(path, sheetName, isEditable);
            return worksheetPart.Worksheet;
        }

        public static Worksheet GetWorksheetBySheetName(this SpreadsheetDocument document, string sheetName)
        {
            var worksheetPart = GetWorksheetPartBySheetName(document, sheetName);
            return worksheetPart.Worksheet;
        }

        public static Row GetRow(this Worksheet worksheet, int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public static Cell GetCell(this Worksheet worksheet, int rowIndex, string columnName)
        {
            var row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
        }

        public static void SetCellValue(this Worksheet worksheet, string columnName, int rowIndex, CellValue cellValue)
        {
            var cell = GetCell(worksheet, rowIndex, columnName);
            cell.CellValue = cellValue;
        }

        public static void SetCellValueAndDataType(this Worksheet worksheet, string columnName, int rowIndex, CellValue cellValue, EnumValue<CellValues> dataType)
        {
            var cell = GetCell(worksheet, rowIndex, columnName);
            cell.CellValue = cellValue;
            cell.DataType = dataType;
        }
    }
}
