using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Domain.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ColumnOrders;

namespace Application.Interfaces
{
   
    public interface IDatabaseRepository
    {
     void AddCalendarEvents(List<CalendarEvent> CalendarEvents);
    void  EditCalendarEvents(List<CalendarEvent> CalendarEvents, List<ExcelRow> ExcelEvents);
    void EditLatestEvents(List<ExcelRow> ExcelEvents);
    int Save();
    }
}

