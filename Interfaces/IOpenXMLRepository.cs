using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Domain.Models;
using static ColumnOrders;

namespace Application.Interfaces
{
    public interface IOpenXMLRepository
    {   MemoryStream OpenDocument(string Path);
        void AddEventsToExcel(MemoryStream fileStream, List<CalendarEvent> CalendarEvents, List<string> EventIDs);
         ExcelRowEvent AddEvent(SheetData sheetData, Row row, CalendarEvent calendarEvent, List<string> EventIDs,  uint RowIndex, ColumnOrder columnOrder);
        
        void EditLatestEvents(MemoryStream fileStream, List<CalendarEvent> CalendarEvents, List<string> EventIDs);
        
        List<CalendarEvent> GetEventsInExcel(string fileName, string sheetName);
        string GetCellValue(Cell cell, WorkbookPart wbPart);
        List<string> GetCellValues(MemoryStream fileStream, string Column);
        List<ExcelRow> GetEventsFromExcel(MemoryStream fileStream,  char[] Columns);

        string GetText(ColumnOrder columnOrder, CalendarEvent Event);
        Cell GetCell(IEnumerable<Cell> Cells, char Column);
        Row GetEditRow(WorkbookPart workbookPart, IEnumerable<Row> rows, string search);
        Row GetRow(SheetData sheetData, uint rowIndex);
        Stylesheet GetStylesheet();

        void UpdateCellValue(WorkbookPart workbookPart, WorksheetPart WorksheetPart, string AddressName, string AddressName2, CalendarEvent Event);
        void UpdateCellValues(string fileName, string sheetName, string Column, List<CalendarEvent> Events);
        string CellValue(WorkbookPart workbookPart, WorksheetPart WorksheetPart, string AddressName);
      
        void InsertCell(WorkbookPart workbookPart, Worksheet worksheet, Row row, uint RowIndex, string Text, char Column);
        Row InsertRow(Worksheet worksheet, uint rowIndex);
        Row InsertRow(SheetData sheetData, uint rowIndex);
        void InsertTextInCell(SpreadsheetDocument spreadsheetDocument, Cell cell, string Text);
        int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart);

        void EditCell(WorkbookPart workbookPart, Worksheet worksheet, Row row, uint RowIndex, string Text, char Column);


      

    }
}
