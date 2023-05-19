using Application.Repository;
using Domain.Models;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class GCalendarEventsController :ControllerBase
    {

        [HttpGet(Name = "GetGEvents")]
        public List<CalendarEvent> Get()
        {
            CalendarEventRepository calendarEventRepository = new CalendarEventRepository();
            //    List<CalendarEvent> DeletedCalendarEvents = GoogleCalendar.Authorize()
            DateTime StartDate = DateTime.Parse("1/1/2020");
            DateTime EndDate = DateTime.Parse("12/31/2023");

            List<CalendarEvent> DeletedCalendarEvents = calendarEventRepository.GetCalendarEvents(StartDate, EndDate);
            List<CalendarEvent> ActiveCalendarEvents = calendarEventRepository.GetCalendarEvents(StartDate, EndDate);
            return ActiveCalendarEvents;

        }
    }
}
