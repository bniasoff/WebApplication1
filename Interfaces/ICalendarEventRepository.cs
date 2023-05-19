using Domain.Models;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;



namespace Application.Interfaces
{
    public interface ICalendarEventRepository
    {
        public List<CalendarEvent> GetCalendarEvents( DateTime StartDate, DateTime EndDate);
        public CalendarService Authorize();
        public Events GetEvents(CalendarService service, String CalendarID, DateTime StartDate, DateTime EndDate);
        public CalendarEvent GetEventDetail(Event CalEvent);
        public Event CreateEvent(CalendarEvent calendarEvent);
        public Event UpdateEvent(Event GoogleEvent, CalendarEvent calendarEvent);
        public List<CalendarEvent> UpdateEvents(List<CalendarEvent> AllEvents, List<CalendarEvent> OldEvents);
    }
}
