using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2010.Excel;
using Google.Apis.Drive.v2.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Storage;
using Microsoft.Extensions.Logging;
using System.Linq;
using System.Linq.Expressions;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using DocumentFormat.OpenXml;
using static ColumnOrders;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Domain.Models;
using Application.Interfaces;

namespace Application.Repository
{

    public class OpenXMLRepository : IOpenXMLRepository
    {



        private string sheetName="Sheet1";

        public MemoryStream OpenDocument(string Path)
        {
            throw new NotImplementedException();
        }



        public ExcelRowEvent AddEvent(SheetData sheetData, Row row, CalendarEvent calendarEvent, List<string> EventIDs, uint RowIndex, ColumnOrder columnOrder)
        {

            //ExcelRowEvent excelRowEvent=new ExcelRowEvent();
            //RowCount += 1;
            //Row row = null;
            // NewRow = false;
            // string cellReference = string.Empty;


            //row,RowIndex,NewRow,text,Column,


            char Column = ';';
            string Location = string.Empty;
            string Address = string.Empty;
            int ColumnIndex = 0;
            bool NewRow = false;
            string Text = null;
            String EventID = calendarEvent.EventID;

            switch (columnOrder.ColumnHead)
            {
                case "EventID":
                    Text = calendarEvent.EventID;
                    EventID = calendarEvent.EventID;
                    if (EventIDs.Contains(Text))
                    { NewRow = false; }

                    if (!EventIDs.Contains(Text))
                    {
                        NewRow = true; ;
                        RowIndex += 1;
                        row = InsertRow(sheetData, RowIndex);
                    }

                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    break;

                case "FamilyName":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.FamilyName;
                    break;
                case "Location":
                    if (calendarEvent.Location != null)
                    {
                        var LocationList = calendarEvent.Location.Split('-').ToList();
                        if (LocationList.Count() > 0)
                        {
                            Location = LocationList[0].Trim();
                            Text = Location;
                        }
                    }

                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    break;
                case "Address":
                    if (calendarEvent.Location != null)
                    {
                        var AddressList = calendarEvent.Location.Split('-').ToList();
                        if (AddressList.Count() > 1)
                        {
                            Address = AddressList[1].Trim();
                            Text = Address;
                        }
                    }

                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    break;
                case "EventDate":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.EventDate.ToOADate().ToString(CultureInfo.InvariantCulture);
                    break;
                case "StartTime":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.StartTime.ToString();
                    break;
                case "EndTime":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column; ;
                    Text = calendarEvent.EndTime.ToString();
                    break;
                case "Phone":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.Phone;
                    break;
                case "Category":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.Category;
                    break;
                case "CreatedDate":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.CreatedDate.ToOADate().ToString(CultureInfo.InvariantCulture);
                    break;
                case "ETagID":
                    ColumnIndex = columnOrder.Order;
                    Column = columnOrder.Column;
                    Text = calendarEvent.ETagID;
                    break;
            }

            var excelRowEvent = new ExcelRowEvent
            {
                Text = Text,
                Column = Column,
                row = row,
                RowIndex = RowIndex,
                NewRow = NewRow
            };
            return excelRowEvent;
        }
        public void AddEventsToExcel(MemoryStream fileStream, List<CalendarEvent> CalendarEvents, List<string> EventIDs)
        {
            Workbook workbook = null;
            Worksheet worksheet = null;
            SheetData sheetData = null;
            WorkbookPart workbookPart = null;
            WorksheetPart worksheetPart = null;

            using (var Document = SpreadsheetDocument.Open("C:\\Pictures\\Excel\\Book1.xlsx", true))

            //using (SpreadsheetDocument Document = SpreadsheetDocument.Open(fileStream, true))
            {

                workbookPart = Document.WorkbookPart;
                workbook = workbookPart.Workbook;

                Sheet sheet = workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
                worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                worksheet = worksheetPart.Worksheet;
                sheetData = worksheet.GetFirstChild<SheetData>();

                var WorkbookStylesPart = workbookPart.WorkbookStylesPart;
                WorkbookStylesPart.Stylesheet = GetStylesheet();
                WorkbookStylesPart.Stylesheet.Save();

                IEnumerable<Row> Rows = sheetData.Elements<Row>();


                uint RowIndex = (uint)Rows.Count();

                ColDict colDic = new ColDict();

                var Today = DateTime.Now;
                var idSet = CalendarEvents.Where(x => x.EventDate > Today).Select(x => x.EventID).ToHashSet();

                int RowCount = 0;

                foreach (CalendarEvent calendarEvent in CalendarEvents)
                {
                    RowCount += 1;
                    Row row = null;
                    // NewRow = false;
                    string cellReferenc  = string.Empty;
                    //String EventID = null;
                    ExcelRowEvent excelRowEvent = null;

                    foreach (ColumnOrder columnOrder in colDic.Columns.Values)
                    {

                        excelRowEvent = AddEvent(sheetData, row, calendarEvent, EventIDs, RowIndex, columnOrder);

                        if (excelRowEvent != null)
                        {

                            if (excelRowEvent.NewRow == true && excelRowEvent.Text != null)
                            {
                                InsertCell(workbookPart, worksheet, excelRowEvent.row, excelRowEvent.RowIndex, excelRowEvent.Text, excelRowEvent.Column);
                            }
                        }

                        //else if (NewRow == false && idSet.Contains(EventID))
                        //{
                        //    EditCell(workbookPart, worksheet, row, RowIndex, Text, Column);
                        //}


                    }
                    if (excelRowEvent.NewRow == true) { sheetData.Append(row); }

                }
                Console.WriteLine(RowCount);
                worksheetPart.Worksheet.Save();
            }
        }



        public string CellValue(WorkbookPart workbookPart, WorksheetPart WorksheetPart, string AddressName)
        {
            string value = null;
            Cell cell = WorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == AddressName).FirstOrDefault();
            if (cell != null)
            {
                if (cell.InnerText.Length > 0)
                {
                    value = cell.InnerText;

                    if (cell.DataType != null)
                    {
                        switch (cell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }

                }
            }
            return value;
        }

        public void EditCell(WorkbookPart workbookPart, Worksheet worksheet, Row row, uint RowIndex, string Text, char Column)
        {
            IEnumerable<Cell> cells = row.Elements<Cell>();
            uint cellCount = (uint)cells.Count();

            Cell cell = null;
            var cellReference = Column + RowIndex.ToString();
            cell = worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();

            ColDict colDic = new ColDict();
            var Cols = ColumnOrders.FillCollumnOrders();

            //char col1 = colDic.Columns.Values.ElementAt(0).Column;
            //char col2 = colDic.Columns.Values.ElementAt(1).Column;
            //char col3 = colDic.Columns.Values.ElementAt(2).Column;
            //char col4 = colDic.Columns.Values.ElementAt(3).Column;
            //char col5 = colDic.Columns.Values.ElementAt(4).Column;
            //char col6 = colDic.Columns.Values.ElementAt(5).Column;
            //char col7 = colDic.Columns.Values.ElementAt(6).Column;
            //char col8 = colDic.Columns.Values.ElementAt(7).Column;
            //char col9 = colDic.Columns.Values.ElementAt(8).Column;
            //char col10 = colDic.Columns.Values.ElementAt(9).Column;
            //char col11 = colDic.Columns.Values.ElementAt(10).Column;

            if (cell == null)
            {
                cell = new Cell();
                cell.CellReference = cellReference;
                cell.CellValue = new CellValue(Text);
                cell.StyleIndex = 1;

                switch (Column)
                {
                    case var value when value == Cols[0]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[1]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[2]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[3]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[4]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        break;
                    case var value when value == Cols[5]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[6]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[7]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[8]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[9]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        break;
                    case var value when value == Cols[10]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;

                }
                row.Append(cell);
            }


            //if (cell == null)
            //{
            //    var newcell = InsertCell(worksheet, row, cellReference, Column, Text);
            //    //InsertTextInCell(spreadsheetDocument, newcell, Text);
            //}
        }
        public void EditLatestEvents(MemoryStream fileStream, List<CalendarEvent> CalendarEvents, List<string> EventIDs)
        {
            Workbook workbook = null;
            Worksheet worksheet = null;
            SheetData sheetData = null;
            WorkbookPart workbookPart = null;
            WorksheetPart worksheetPart = null;

            //using (SpreadsheetDocument Document = SpreadsheetDocument.Open("C:\\Pictures\\Excel\\Photo Events.xlsx", true))

            using (SpreadsheetDocument Document = SpreadsheetDocument.Open(fileStream, true))
            {

                ColDict colDic = new ColDict();

                workbookPart = Document.WorkbookPart;
                workbook = workbookPart.Workbook;

                Sheet sheet = workbook.Descendants<Sheet>().Where(s => s.Name == "Events").FirstOrDefault();
                worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                worksheet = worksheetPart.Worksheet;
                sheetData = worksheet.GetFirstChild<SheetData>();

                SharedStringTablePart shareStringPart;
                shareStringPart = Document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

                IEnumerable<Row> Rows = sheetData.Elements<Row>();
                uint RowIndex = (uint)Rows.Count();


                var CalEventsToEdit = CalendarEvents.Where(ev => ev.EventDate >= DateTime.Now.AddDays(-1)).ToList();
                //var CalEventsToEdit = CalendarEvents.Where(ev => ev.FamilyName == "Applebaum").ToList();


                foreach (CalendarEvent CalendarEvent in CalEventsToEdit)
                {
                    var EventID = CalendarEvent.EventID;
                    var EditRow = GetEditRow(workbookPart, Rows, EventID);
                    if (EditRow != null)
                    {
                        uint rowIndex = EditRow.RowIndex;

                        IEnumerable<Cell> EditCells = EditRow.Descendants<Cell>();

                        foreach (ColumnOrder columnOrder in colDic.Columns.Values)
                        {
                            string Text = GetText(columnOrder, CalendarEvent);
                            int index = InsertSharedStringItem(Text, shareStringPart);

                            string cellReference = columnOrder.Column + rowIndex.ToString();

                            Cell cell = EditCells.Where(c => c.CellReference == cellReference).FirstOrDefault();

                            if (cell == null)
                            {
                                Cell refCell = null;
                                foreach (Cell cell2 in EditRow.Elements<Cell>())
                                {
                                    if (string.Compare(cell2.CellReference.Value, cellReference, true) > 0)
                                    {
                                        refCell = cell2;
                                        break;
                                    }
                                }

                                cell = new Cell() { CellReference = cellReference };
                                EditRow.InsertBefore(cell, refCell);
                                worksheet.Save();
                                cell.CellValue = new CellValue(index.ToString());
                                cell.StyleIndex = 1;
                                switch (columnOrder.ColumnHead)
                                {
                                    case String value when value == "EventDate":
                                        //cell.CellValue = new CellValue(Text);
                                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                        break;
                                    case String value when value == "CreatedDate":
                                        // cell.CellValue = new CellValue(Text);
                                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                        break;
                                    default:
                                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                                        break;
                                }
                            }

                            if (cell != null)
                            {
                                cell.StyleIndex = 1;
                                switch (columnOrder.ColumnHead)
                                {
                                    case String value when value == "EventDate":
                                        cell.CellValue = new CellValue(Text);
                                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                        break;
                                    case String value when value == "CreatedDate":
                                        cell.CellValue = new CellValue(Text);
                                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                        break;
                                    default:
                                        cell.CellValue = new CellValue(index.ToString());
                                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                                        break;
                                }
                            }
                        }
                    }
                }
                worksheetPart.Worksheet.Save();
            }
        }



        public Cell GetCell(IEnumerable<Cell> Cells, char Column)
        {
            Cell cell = null;
            cell = Cells.Where(c => c.CellReference.Value.FirstOrDefault() == Column).FirstOrDefault();

            // return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();

            //    switch (Column)
            //    {
            //        case var value when value == col1:

            //            break;
            //        case var value when value == col2:

            //            break;
            //        case var value when value == col3:

            //            break;
            //        case var value when value == col4:

            //            break;
            //        case var value when value == col5:

            //            break;
            //        case var value when value == col6:

            //            break;
            //        case var value when value == col7:

            //            break;
            //    }

            return cell;
        }

        public string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            string value = null;
            if (cell != null)
            {
                value = cell.InnerText;

                // If the cell represents an integer number, you are done.
                // For dates, this code returns the serialized value that
                // represents the date. The code handles strings and
                // Booleans individually. For shared strings, the code
                // looks up the corresponding value in the shared string
                // table. For Booleans, the code converts the value into
                // the words TRUE or FALSE.
                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in
                            // the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(
                                     int.Parse(value)).InnerText;
                            }

                            break;
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value =

                                    "FALSE";
                                    break;
                                default:
                                    value =

                                    "TRUE";
                                    break;
                            }
                            break;
                    }
                }
                else
                {

                    string Column = cell.CellReference;

                    bool Converted = false;
                    switch (Column.Substring(0, 1))
                    {
                        case "E":
                            double DateNumber;
                            Converted = double.TryParse(value, out DateNumber);
                            value = DateTime.FromOADate(DateNumber).ToString();
                            break;
                        case "F":
                            double StartDateNumber;
                            Converted = double.TryParse(value, out StartDateNumber);
                            value = DateTime.FromOADate(StartDateNumber).ToString();
                            break;
                        case "G":
                            double EndDateNumber;
                            Converted = double.TryParse(value, out EndDateNumber);
                            value = DateTime.FromOADate(EndDateNumber).ToString();
                            break;

                    }

                }
            }
            return value;
        }

        public List<string> GetCellValues(MemoryStream fileStream, string Column)
        {
             using (SpreadsheetDocument Document = SpreadsheetDocument.Open("C:\\Pictures\\Excel\\Book1.xlsx", true))
            //using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
           // using (SpreadsheetDocument Document = SpreadsheetDocument.Open(fileStream, false))
            {

                WorkbookPart workbookPart = Document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                //if (sheet == null)
                //{
                //    throw new ArgumentException("Events");
                //}

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                IEnumerable<Row> rows = sheetData.Elements<Row>();
                uint RowCount = (uint)rows.Count();
                string AddressName = null;

                var EventIDs = new List<string>();

                for (uint i = 2; i <= RowCount; i++)
                {
                    AddressName = Column + i;
                    string ColumnValue = CellValue(workbookPart, worksheetPart, AddressName);
                    AddressName = null;
                    if (!string.IsNullOrWhiteSpace(ColumnValue))
                    {
                        if (ColumnValue.Length > 10)
                            EventIDs.Add(ColumnValue);
                    }

                }
                return EventIDs;
            }
        }

        public List<ExcelRow> GetEventsFromExcel(MemoryStream fileStream, char[] Columns)
        {
            List<ExcelRow> Events = new List<ExcelRow>();
             using (SpreadsheetDocument Document = SpreadsheetDocument.Open("C:\\Pictures\\Excel\\Book1.xlsx", true))
          // using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
           // using (SpreadsheetDocument Document = SpreadsheetDocument.Open(fileStream, true))
            {

                WorkbookPart workbookPart = Document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();



                if (sheet == null)
                {
                    throw new ArgumentException("Events");
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                IEnumerable<Row> rows = sheetData.Elements<Row>();
                uint RowCount = (uint)rows.Count();
                string AddressName = null;


                ColDict colDic = new ColDict();

                var EventIDs = new List<string>();

                Worksheet worksheet = null;
                worksheet = worksheetPart.Worksheet;
                SheetData Rows = (SheetData)worksheet.ChildElements.GetItem(4);
                //int r = 1;
                // foreach (Row row in Rows.Skip(1))
                //  {

                var WorkbookStylesPart = workbookPart.WorkbookStylesPart;
                WorkbookStylesPart.Stylesheet = GetStylesheet();
                WorkbookStylesPart.Stylesheet.Save();

                for (uint r = 2; r <= RowCount; r++)
                {
                    // r = r += 1;
                    string EventIDCol = "A" + r.ToString();
                    string EventID = CellValue(workbookPart, worksheetPart, EventIDCol);


                    ExcelRow excelrow = new ExcelRow();


                    foreach (ColumnOrder columnOrder in colDic.Columns.Values)
                    {

                        //char Column = columnOrder.Column;
                        AddressName = columnOrder.Column.ToString() + r;
                        string ColumnValue = CellValue(workbookPart, worksheetPart, AddressName);

                        if (!string.IsNullOrWhiteSpace(ColumnValue))
                        {
                            switch (columnOrder.ColumnHead)
                            {
                                case "EventID":
                                    excelrow.EventID = ColumnValue;
                                    break;
                                case "FamilyName":
                                    excelrow.FamilyName = ColumnValue;
                                    break;
                                case "Location":
                                    excelrow.Location = ColumnValue;
                                    break;
                                case "Address":
                                    excelrow.Address = ColumnValue;
                                    break;
                                case "EventDate":
                                    double OutVal;
                                    bool Parsed = double.TryParse(ColumnValue, out OutVal);
                                    //if (double.TryParse(ColumnValue.Replace("..", "."), out OutVal))
                                    //if (double.TryParse(ColumnValue, NumberStyles.Any, new CultureInfo("de-DE"), out OutVal))
                                    //Console.WriteLine(result);

                                    if (Parsed)
                                    {
                                        DateTime EventDate;
                                        EventDate = DateTime.FromOADate(OutVal);
                                        excelrow.EventDate = EventDate;
                                    }
                                    break;
                                case "StartTime":
                                    TimeSpan StartTime;
                                    bool TimeSpanParsed = TimeSpan.TryParse(ColumnValue, out StartTime);
                                    if (TimeSpanParsed) { excelrow.StartTime = StartTime; }
                                    break;
                                case "EndTime":
                                    TimeSpan EndTime;
                                    bool TimeSpanParsed2 = TimeSpan.TryParse(ColumnValue, out EndTime);
                                    if (TimeSpanParsed2) { excelrow.EndTime = EndTime; }
                                    break;
                                case "Phone":
                                    excelrow.Phone = ColumnValue;
                                    break;
                                case "Category":
                                    excelrow.Category = ColumnValue;
                                    break;
                                case "ChargeAmount":
                                    decimal ChargeAmount;
                                    bool decimalParsed = Decimal.TryParse(ColumnValue, out ChargeAmount);
                                    if (decimalParsed) { excelrow.ChargeAmount = ChargeAmount; }
                                    break;
                                case "PaidAmount":
                                    decimal PaidAmount;
                                    bool decimalParsed2 = Decimal.TryParse(ColumnValue, out PaidAmount);
                                    if (decimalParsed2) { excelrow.PaidAmount = PaidAmount; }
                                    break;
                                case "Balence":

                                    excelrow.Balence = ColumnValue;
                                    break;
                                case "ToDo":
                                    bool ToDo;
                                    ToDo = ColumnValue.Equals("1.0");
                                    excelrow.ToDo = ToDo;
                                    break;
                                case "Ready":
                                    bool Ready;
                                    Ready = ColumnValue.Equals("1.0");
                                    excelrow.Ready = Ready;
                                    break;
                                case "Sent":
                                    bool Sent;
                                    Sent = ColumnValue.Equals("1.0");
                                    excelrow.Sent = Sent;
                                    break;
                                case "Paid":
                                    bool Paid;
                                    //int Paid2;
                                    Paid = ColumnValue.Equals("1.0");
                                    excelrow.Paid = Paid;
                                    break;
                                case "Referred":
                                    excelrow.Referred = ColumnValue;
                                    break;
                            }
                        }
                    }
                    if (excelrow.EventID != null) { Events.Add(excelrow); }
                }
            }
            return Events;
        }
        public Row GetEditRow(WorkbookPart workbookPart, IEnumerable<Row> rows, string search)
        {
            string index = string.Empty;
            string value = null;
            uint? indexval = null;

            foreach (var rowItem in rows)
            {


                //var Cells = rowItem.getChildElements<Cell>;

                Cell FirstCell = rowItem.GetFirstChild<Cell>();
                if (FirstCell != null)
                {
                    CellValue CellRef = FirstCell.GetFirstChild<CellValue>();
                    if (CellRef != null)
                    {
                        var FirstLetter = FirstCell.CellReference.ToString().FirstOrDefault();
                        if (FirstLetter == 'A')
                        {
                            int intVal = -1;
                            var IsInt = int.TryParse(CellRef.InnerText, out intVal);

                            if (IsInt)
                            {
                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                value = stringTable.SharedStringTable.ElementAt(intVal).InnerText;
                            }
                            if (!IsInt) { value = search; }


                            bool isFound = value.Trim().ToLower().Contains(search.Trim().ToLower());

                            if (isFound)
                            {
                                // index = $"[{rowItem.RowIndex}, {GetColumnIndex(FirstCell.CellReference)}]";
                                indexval = rowItem.RowIndex.Value;
                                // return indexval;
                            }

                            if (isFound)
                            {
                                return rowItem;
                            }
                        }
                    }
                }
            }
            return null;
        }

        public List<CalendarEvent> GetEventsInExcel(string fileName, string sheetName)
        {
            List<CalendarEvent> Events = new List<CalendarEvent>();
            Workbook workbook = null;
            Worksheet worksheet = null;
            SheetData sheetData = null;
            SharedStringItem item = null;
            WorkbookPart workbookPart = null;
            WorksheetPart worksheetPart = null;
            string sheetId = string.Empty;
            ;

            using (SpreadsheetDocument Document = SpreadsheetDocument.Open(fileName, false))

            {
                workbookPart = Document.WorkbookPart;
                workbook = workbookPart.Workbook;

                //var WorkbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                //WorkbookStylesPart.Stylesheet = Class4.GetStylesheet();
                //WorkbookStylesPart.Stylesheet.Save();

                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Events").FirstOrDefault();
                sheetId = sheet.Id;
                worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
                worksheet = worksheetPart.Worksheet;
                sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                SheetData Rows = (SheetData)worksheet.ChildElements.GetItem(4);

                string parseCellValue;

                foreach (Row row in Rows.Skip(1))
                {
                    CalendarEvent Event = new CalendarEvent();

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        parseCellValue = GetCellValue(cell, workbookPart);
                        string Column = cell.CellReference.Value.Substring(0, 1);
                        string eventDate = string.Empty; string createdDate = string.Empty;
                        DateTime parsedDate;
                        TimeSpan parsedTime;
                        switch (Column)
                        {
                            case "A":
                                Event.EventID = parseCellValue;
                                break;
                            case "B":
                                Event.FamilyName = parseCellValue;
                                break;
                            case "C":
                                Event.Location = parseCellValue;
                                break;
                            case "D":
                                Event.Address = parseCellValue;
                                break;
                            case "K":
                                parsedDate = DateTime.Parse(parseCellValue);
                                Event.EventDate = parsedDate;
                                break;
                            case "L":
                                if (Event.EventDate != null)
                                { eventDate = Event.EventDate.ToString("MM/dd/yyyy"); }
                                parseCellValue = parseCellValue.Replace("12/30/1899", eventDate);
                                parseCellValue = eventDate.ToString() + " " + parseCellValue;

                                parsedDate = DateTime.Parse(parseCellValue);
                                parsedTime = TimeSpan.FromTicks(parsedDate.Ticks);
                                Event.StartTime = parsedTime;
                                break;
                            case "M":
                                if (Event.EventDate != null)
                                { eventDate = Event.EventDate.ToString("MM/dd/yyyy"); }
                                parseCellValue = parseCellValue.Replace("12/30/1899", eventDate);
                                parseCellValue = eventDate.ToString() + " " + parseCellValue;

                                parsedDate = DateTime.Parse(parseCellValue);
                                parsedTime = TimeSpan.FromTicks(parsedDate.Ticks);
                                Event.EndTime = parsedTime;
                                break;
                            case "N":
                                Event.Phone = parseCellValue;
                                break;
                            case "O":
                                Event.Category = parseCellValue;
                                break;
                            case "Q":
                                if (Event.CreatedDate != null)
                                { createdDate = Event.CreatedDate.ToString("MM/dd/yyyy"); }
                                parseCellValue = parseCellValue.Replace("12/30/1899", createdDate);
                                parseCellValue = createdDate.ToString() + " " + parseCellValue;

                                parsedDate = DateTime.Parse(parseCellValue);
                                Event.CreatedDate = parsedDate;
                                break;
                            case "R":
                                Event.ETagID = parseCellValue;
                                break;

                        }

                    }
                    Events.Add(Event);
                }


            }

            return Events;
        }

        public Row GetRow(SheetData sheetData, uint rowIndex)
        {
            var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                sheetData.Append(row);
            }
            return row;
        }

        //public int GetStylesheet2(SpreadsheetDocument spreadsheetDoc)
        //{


        //    // get the stylesheet from the current sheet    
        //    var stylesheet = spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet;
        //    // cell formats are stored in the stylesheet's NumberingFormats
        //    var numberingFormats = stylesheet.NumberingFormats;

        //    // cell format string               
        //    const string dateFormatCode = "dd/mm/yyyy";
        //    // first check if we find an existing NumberingFormat with the desired formatcode
        //    var dateFormat = numberingFormats.OfType<NumberingFormat>().FirstOrDefault(format => format.FormatCode == dateFormatCode);
        //    // if not: create it
        //    if (dateFormat == null)
        //    {
        //        dateFormat = new NumberingFormat
        //        {
        //            NumberFormatId = UInt32Value.FromUInt32(164),  // Built-in number formats are numbered 0 - 163. Custom formats must start at 164.
        //            FormatCode = StringValue.FromString(dateFormatCode)
        //        };
        //        numberingFormats.AppendChild(dateFormat);
        //        // we have to increase the count attribute manually ?!?
        //        numberingFormats.Count = Convert.ToUInt32(numberingFormats.Count());
        //        // save the new NumberFormat in the stylesheet
        //        stylesheet.Save();
        //    }
        //    // get the (1-based) index of the dateformat
        //    var dateStyleIndex = numberingFormats.ToList().IndexOf(dateFormat) + 1;
        //    return dateStyleIndex;
        //}


        public Stylesheet GetStylesheet()
        {


            var StyleSheet = new Stylesheet();

            // Create "fonts" node.
            var Fonts = new Fonts();
            Fonts.Append(new Font()
            {
                FontName = new FontName() { Val = "Calibri" },
                FontSize = new FontSize() { Val = 11 },
                FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
            });

            Fonts.Count = (uint)Fonts.ChildElements.Count;

            // Create "fills" node.
            var Fills = new Fills();
            Fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            });
            Fills.Append(new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
            });

            Fills.Count = (uint)Fills.ChildElements.Count;

            // Create "borders" node.
            var Borders = new Borders();
            Borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });

            Borders.Count = (uint)Borders.ChildElements.Count;

            // Create "cellStyleXfs" node.
            var CellStyleFormats = new CellStyleFormats();
            CellStyleFormats.Append(new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });

            CellStyleFormats.Count = (uint)CellStyleFormats.ChildElements.Count;

            // Create "cellXfs" node.
            var CellFormats = new CellFormats();

            // A default style that works for everything but DateTime
            CellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            });

            // A style that works for DateTime (just the date)
            CellFormats.Append(new CellFormat()
            {
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 14, // or 22 to include the time
                FormatId = 0,
                ApplyNumberFormat = true
            });

            CellFormats.Count = (uint)CellFormats.ChildElements.Count;

            // Create "cellStyles" node.
            var CellStyles = new CellStyles();
            CellStyles.Append(new CellStyle()
            {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0
            });
            CellStyles.Count = (uint)CellStyles.ChildElements.Count;

            // Append all nodes in order.
            StyleSheet.Append(Fonts);
            StyleSheet.Append(Fills);
            StyleSheet.Append(Borders);
            StyleSheet.Append(CellStyleFormats);
            StyleSheet.Append(CellFormats);
            StyleSheet.Append(CellStyles);

            return StyleSheet;
        }

        public string GetText(ColumnOrder columnOrder, CalendarEvent Event)
        {
            var Cols = ColumnOrders.FillCollumnOrders();


            // ColDict colDic = new ColDict();
            string Text = string.Empty;


            switch (columnOrder.Column)
            {
                case var value when value == Cols[0]:
                    Text = Event.EventID;
                    break;
                case var value when value == Cols[1]:
                    Text = Event.FamilyName;
                    break;
                case var value when value == Cols[2]:
                    if (Event.Location != null)
                    {
                        var LocationList = Event.Location.Split('-').ToList();

                        if (LocationList.Count() > 0)
                        {
                            Text = LocationList[0].Trim();

                        }
                    }
                    //Text = Event.Location;
                    break;
                case var value when value == Cols[3]:
                    //if (Event.Location != null)

                    //{
                    //    var AddressList = Event.Location.Split('-').ToList();
                    //    if (AddressList.Count() > 1)
                    //    {
                    //        Text = AddressList[1].Trim();

                    //    }
                    //}
                    Text = Event.Address;
                    break;
                case var value when value == Cols[11]:
                    Text = Event.EventDate.ToOADate().ToString();
                    break;
                case var value when value == Cols[12]:
                    Text = Event.StartTime.ToString();
                    break;
                case var value when value == Cols[13]:
                    Text = Event.EndTime.ToString();
                    break;
                case var value when value == Cols[14]:
                    Text = Event.Phone;
                    break;
                case var value when value == Cols[15]:
                    Text = Event.Category;
                    break;
                case var value when value == Cols[17]:
                    Text = Event.CreatedDate.ToOADate().ToString();
                    break;
                case var value when value == Cols[18]:
                    Text = Event.ETagID;
                    break;
            }

            return Text;
        }

        public void InsertCell(WorkbookPart workbookPart, Worksheet worksheet, Row row, uint RowIndex, string Text, char Column)
        {
            IEnumerable<Cell> cells = row.Elements<Cell>();
            uint cellCount = (uint)cells.Count();

            Cell cell = null;
            var cellReference = Column + RowIndex.ToString();
            cell = worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();

            ColDict colDic = new ColDict();
            var Cols = ColumnOrders.FillCollumnOrders();


            if (cell == null)
            {
                cell = new Cell();
                cell.CellReference = cellReference;
                cell.CellValue = new CellValue(Text);
                cell.StyleIndex = 1;

                switch (Column)
                {
                    case var value when value == Cols[0]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[1]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[2]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[3]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[4]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[5]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[6]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[7]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[8]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[9]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[11]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        break;
                    case var value when value == Cols[12]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[13]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[14]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[15]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[16]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case var value when value == Cols[17]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        break;
                    case var value when value == Cols[18]:
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                }
                row.Append(cell);
            }


            //if (cell == null)
            //{
            //    var newcell = InsertCell(worksheet, row, cellReference, Column, Text);
            //    //InsertTextInCell(spreadsheetDocument, newcell, Text);
            //}
        }

        public Row InsertRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = null;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            // Check if the worksheet contains a row with the specified row index.
            row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            return row;
        }

        public Row InsertRow(SheetData sheetData, uint rowIndex)
        {
            Row row = null;
            row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
            }
            return row;
        }

        public int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        public void InsertTextInCell(SpreadsheetDocument spreadsheetDocument, Cell cell, string Text)
        {
            SharedStringTablePart shareStringPart;
            if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            int index = InsertSharedStringItem(Text, shareStringPart);
            cell.CellValue = new CellValue(index.ToString());
        }

        public void UpdateCellValue(WorkbookPart workbookPart, WorksheetPart WorksheetPart, string AddressName, string AddressName2, CalendarEvent Event)
        {
            string value = null;
            Cell cell = WorksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == AddressName2).FirstOrDefault();
            if (cell != null)
            {
                if (cell.InnerText.Length > 0)
                {
                    //cell.= Event.ID;

                    if (cell.DataType != null)
                    {
                        switch (cell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                if (stringTable != null)
                                {
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }

                }
            }
        }

        public void UpdateCellValues(string fileName, string sheetName, string Column, List<CalendarEvent> Events)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {

                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                if (sheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                IEnumerable<Row> rows = sheetData.Elements<Row>();
                uint RowCount = (uint)rows.Count();
                string AddressName = null;

                uint RowCount2 = 2;
                foreach (var row in rows)
                {

                    AddressName = Column + RowCount2;
                    RowCount2 = RowCount2 + 1;
                    string ColumnValue = CellValue(workbookPart, worksheetPart, AddressName);

                    if (!string.IsNullOrWhiteSpace(ColumnValue))
                    {
                        string ETag = ColumnValue;
                        var CalendarEvent = Events.Where(x => x.EventID == ETag).FirstOrDefault();
                        if (CalendarEvent != null)
                        {

                            if (CalendarEvent.EventID != null)
                            {
                                string cellReference = "K" + RowCount2;


                                var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);

                                if (cell == null)
                                {

                                    //var newcell = InsertCell(worksheetPart.Worksheet, row, (uint)RowCount2,  15, Column);
                                    //InsertTextInCell(document, newcell, CalendarEvent.ID);


                                    //cell = new Cell() { CellReference = cellReference };


                                    ////cell.CellReference   = cellReference ;
                                    //Cell refCell = null;
                                    //var newChild = row.InsertBefore(cell, refCell);
                                    //InsertTextInCell(document, newChild, CalendarEvent.ID);
                                    //worksheetPart.Worksheet.Save();
                                    ////worksheet.Save();

                                }

                            }

                        }
                    }
                }

                worksheetPart.Worksheet.Save();
                document.Close();



            }
        }


    }

}

















