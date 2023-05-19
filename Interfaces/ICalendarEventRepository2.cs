//using Domain.Models;
//using Google.Apis.Calendar.v3;
//using Google.Apis.Calendar.v3.Data;



//namespace Application.Interfaces
//{
//    public interface ICalendarEventRepository
//    {
//        ICollection<CalendarEvent> GetCalendarEvent();
//        CalendarEvent GetCalendarEvent(string EventID);
//        // ICollection<CalendarEvent> GetCalendarEventByCategory(int EventID);
//        bool CalendarEventExists(string EventID);
//        bool CreateCalendarEvent(CalendarEvent calendarEvent);
//        bool UpdateCalendarEvent(CalendarEvent calendarEvent);
//        bool DeleteCalendarEvent(CalendarEvent calendarEvent);
//        CalendarService Authorize();
//        Events GetEvents(CalendarService service);
//        Event CreateEvent(CalendarEvent calendarEvent);
//        Event UpdateEvent(Event _event, CalendarEvent calendarEvent);
//        List<CalendarEvent> UpdateEvents(List<CalendarEvent> AllEvents, List<CalendarEvent> OldEvents);
//        CalendarEvent GetEventDetail(Event _event);
//        List<CalendarEvent> GetEventDetails(Events events);
//        List<CalendarEvent> GetCalendarEvents();

//        bool Save();
//    }
//}
