
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Text.RegularExpressions;
using Application.Interfaces;
using Domain.Models;
using Application.Helper;
using System.Globalization;

namespace Application.Repository
{

    public class CalendarEventRepository : ICalendarEventRepository
    {
        static string[] Scopes = { CalendarService.Scope.Calendar, CalendarService.Scope.CalendarEvents };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
       
        public string CalendarID = "3dc3debda4eb8d023e70ae94e7043de683c6af12080c8a7515bd4e2101625b4c@group.calendar.google.com";
        public DateTime StartDate;
        public DateTime EndDate;

        public List<CalendarEvent> GetCalendarEvents(DateTime StartDate, DateTime EndDate)
        {
            CalendarService Service = Authorize();
            Events events = GetEvents(Service, CalendarID, StartDate, EndDate);

            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();

            if (events != null)
            {
                foreach (Event eventItem in events.Items)
                {
                    CalendarEvent calendarEvent = GetEventDetail(eventItem);
                    CalendarEvents.Add(calendarEvent);
                }
            }

            return CalendarEvents;
        }

        public CalendarService Authorize()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token2.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }


        public Events GetEvents(CalendarService service, String CalendarID, DateTime StartDate, DateTime EndDate)
        {
            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();

            EventsResource.ListRequest request = service.Events.List(CalendarID);
            request.TimeMin = DateTime.Parse(StartDate.ToString()); ;
            request.TimeMax = DateTime.Parse(EndDate.ToString()); ;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            //request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            Events events = request.Execute();

            return events;
        }

        public CalendarEvent GetEventDetail(Event CalEvent)
        {

            DateTime EventDate = new DateTime();
            TimeSpan? StartTime = null;
            TimeSpan? EndTime = null;
            DateTime CreatedDate = new DateTime();

            string Phone = string.Empty, FamilyName = string.Empty, Category = string.Empty, Address = string.Empty;

            var EventNameList = CalEvent.Summary?.Split('-').ToList();
            var LocationList = CalEvent.Location?.Split('-').ToList();

            if (CalEvent.Start.DateTime != null)
            {
                EventDate = CalEvent.Start.DateTime.Value.Date;
                StartTime = CalEvent.Start.DateTime.Value.TimeOfDay;
            }

            if (CalEvent.End.DateTime != null)
            {
                EndTime = CalEvent.End.DateTime.Value.TimeOfDay;
            }

            if (CalEvent.Created.Value != null)
            {
                CreatedDate = CalEvent.Created.Value;

            }

            if (CalEvent.Description != null)
            {
                Phone = CalEvent.Description.Replace("Phone:", "").Replace("\n", "").Replace("\r", "").Trim();

            }

            if (EventNameList != null)
            {
                if (EventNameList.Count > 0)
                {
                    Category = EventNameList[0].Trim();
                }
                if (EventNameList.Count > 1)
                {
                    //FamilyName = EventNameList[1].Trim();
                    FamilyName = EventNameList[1].Replace("\n", "").Replace("\r", "").Trim();
                }
            }

            if (LocationList != null)
            {
                if (LocationList.Count() > 1)
                {
                    // Location= LocationList[0].Trim();
                    Address = LocationList[1].Trim();
                }
            }
            var ETag = Regex.Replace(CalEvent.ETag, @"[\W_]", "");

            CalendarEvent calendarEvent = new CalendarEvent
            {
                //CalendarName = events.Summary,              
                EventID = CalEvent.Id,
                EventDate = EventDate,
                StartTime = StartTime,
                EndTime = EndTime,
                FamilyName = FamilyName,
                Location = CalEvent.Location,
                Address = Address,
                Phone = Phone,
                Category = Category,
                CreatedDate = CreatedDate,
                ETagID = ETag
            };
            return calendarEvent;
        }

        public Event CreateEvent(CalendarEvent calendarEvent)
        {

            DateTime StartTime = DateTime.Parse(calendarEvent.EventDate.ToString());
            StartTime = StartTime.AddMinutes(calendarEvent.StartTime.Value.TotalMinutes);

            DateTime EndTime = DateTime.Parse(calendarEvent.EventDate.ToString());
            EndTime = EndTime.AddMinutes(calendarEvent.EndTime.Value.TotalMinutes);

            Event newEvent = new Event()
            {
                Description = $"Phone - {calendarEvent.Phone}",
                Summary = $"{calendarEvent?.Category} - {calendarEvent?.FamilyName}",
                Location = $"{calendarEvent?.Location} - {calendarEvent?.Address}",

                Start = new EventDateTime() { DateTime = StartTime, TimeZone = "America/New_York" },
                End = new EventDateTime() { DateTime = EndTime, TimeZone = "America/New_York" }
            };
            return newEvent;
        }


        public Event UpdateEvent(Event GoogleEvent, CalendarEvent calendarEvent)
        {
            GoogleEvent.Description = $"Phone - {calendarEvent.Phone}";
            GoogleEvent.Summary = $"{calendarEvent?.Category} - {calendarEvent?.FamilyName}";
            GoogleEvent.Location = $"{calendarEvent?.Location} - {calendarEvent?.Address}";

            DateTime StartTime = DateTime.Parse(calendarEvent.EventDate.ToString());
            StartTime = StartTime.AddMinutes(calendarEvent.StartTime.Value.TotalMinutes);

            DateTime EndTime = DateTime.Parse(calendarEvent.EventDate.ToString());
            EndTime = EndTime.AddMinutes(calendarEvent.EndTime.Value.TotalMinutes);

            GoogleEvent.Start = new EventDateTime() { DateTime = StartTime, TimeZone = "America/New_York" };
            GoogleEvent.End = new EventDateTime() { DateTime = EndTime, TimeZone = "America/New_York" };

            return GoogleEvent;
        }


        public List<CalendarEvent> UpdateEvents(List<CalendarEvent> AllEvents, List<CalendarEvent> OldEvents)
        {
            List<CalendarEvent> NewEvents = new List<CalendarEvent>();
            var AllEventID = AllEvents.Select(x => x.EventID).ToList();
            var OldEventID = OldEvents.Select(x => x.EventID).ToList();

            var result = from a in AllEventID where !OldEventID.Contains(a) select a.ToString();

            foreach (string item in result)
            {
                CalendarEvent? NewEvent = null;
                NewEvent = AllEvents.FirstOrDefault(x => x.EventID == item);
                if (NewEvent != null) { NewEvents.Add(NewEvent); }
            }


            foreach (CalendarEvent EventItem in AllEvents.Where(x => x.EventDate >= DateTime.Now))
            {
                CalendarEvent? DataEvent = OldEvents.Where(x => x.EventID == EventItem.EventID).FirstOrDefault();
                if (DataEvent != null)
                {
                    DataEvent.FamilyName = EventItem.FamilyName;
                    DataEvent.Location = EventItem.Location;
                    DataEvent.Address = EventItem.Address;
                    DataEvent.Phone = EventItem.Phone;
                    DataEvent.StartTime = EventItem.StartTime;
                    DataEvent.EndTime = EventItem.EndTime;
                    DataEvent.ETagID = EventItem.ETagID;
                }
            }


            return NewEvents;
        }

    }
}

















