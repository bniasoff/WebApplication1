using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using static ColumnOrders;
using DocumentFormat.OpenXml.Spreadsheet;
using Application.Interfaces;
using Domain.Models;
using Persistence;

namespace Application.Repository
{
    public class DatabaseRepository : IDatabaseRepository
    {
        //private readonly ILogger<DatabaseRepository> _log;
        private readonly IConfiguration _config;
        private DataContext _context;
        public DatabaseRepository(IConfiguration config, DataContext dataContext)
        {
            //_log = log;
            _config = config;
            _context = dataContext;

        }

        public void AddCalendarEvents(List<CalendarEvent> CalendarEvents)
        {

            _context.CalendarEvents.AddRange(CalendarEvents);
        }

        public void EditCalendarEvents(List<CalendarEvent> CalendarEvents, List<ColumnOrders.ExcelRow> ExcelEvents)
        {
            //var CalendarEvent2 = ExcelEvents.Where(e => e.ChargeAmount > 0).ToList();

            foreach (ExcelRow row in ExcelEvents)
            {
                var CalendarEvent = _context.CalendarEvents.Where(e => e.EventID == row.EventID).FirstOrDefault();
                if (CalendarEvent != null)
                {
                    CalendarEvent.Charge = row.ChargeAmount;
                    CalendarEvent.Paid = row.PaidAmount;
                    CalendarEvent.ToDo = row.ToDo;
                    CalendarEvent.Ready = row.Ready;
                    CalendarEvent.Sent = row.Sent;
                    CalendarEvent.Referred = row.Referred;
                }

            }
        }

        public void EditLatestEvents(List<ExcelRow> ExcelEvents)
        {
            List<ExcelRow> ExcelEventsToEdit = ExcelEvents.Where(e => e.EventDate >= DateTime.Now.AddDays(-1)).ToList();

            foreach (ExcelRow row in ExcelEventsToEdit)
            {
                var CalendarEvent = _context.CalendarEvents.Where(e => e.EventID == row.EventID).FirstOrDefault();
                if (CalendarEvent != null)
                {
                    CalendarEvent.FamilyName = row.FamilyName;

                    CalendarEvent.Location = row.Location;
                    CalendarEvent.Address = row.Address;

                    CalendarEvent.EventDate = row.EventDate;
                    CalendarEvent.StartTime = row.StartTime;
                    CalendarEvent.EndTime = row.EndTime;
                    CalendarEvent.Phone = row.Phone;
                    CalendarEvent.Category = row.Category;   
                   
                    CalendarEvent.Charge = row.ChargeAmount;
                    CalendarEvent.Paid = row.PaidAmount;
                    CalendarEvent.ToDo = row.ToDo;
                    CalendarEvent.Ready = row.Ready;
                    CalendarEvent.Sent = row.Sent;
                    CalendarEvent.Referred = row.Referred;
                }

            }
        }

        public int Save()
        {
            int RecordsUpdated=_context.SaveChanges();
            return RecordsUpdated;
        }
    }

}

















