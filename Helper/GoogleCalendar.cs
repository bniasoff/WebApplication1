
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Domain.Models;

namespace Application.Helper

{

    public static class GoogleCalendar
    {
        static string[] Scopes = { CalendarService.Scope.Calendar, CalendarService.Scope.CalendarEvents };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        public static List<CalendarEvent> GetCalendarEvents()
        {
            var Service = Authorize();
            var Events = GetEvents(Service);
            var CalendarEvents = GetEventDetails(Events);

            return CalendarEvents;
        }


        public static CalendarService Authorize()
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

        public static Events GetEvents(CalendarService service)
        {

            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();

            EventsResource.ListRequest request = service.Events.List("3dc3debda4eb8d023e70ae94e7043de683c6af12080c8a7515bd4e2101625b4c@group.calendar.google.com");
            request.TimeMin = DateTime.Parse("1/1/2020"); ;
            request.TimeMax = DateTime.Parse("12/31/2023"); ;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            //request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            Events events = request.Execute();

            return events;
        }


        public static List<CalendarEvent> UpdateEvents(List<CalendarEvent> AllEvents, List<CalendarEvent> OldEvents)
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

        public static CalendarEvent GetEventDetail(Event CalEvent)
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

        public static List<CalendarEvent> GetEventDetails(Events events)
        {
            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();
            if (events.Items != null && events.Items.Count > 0)
            {

                foreach (var eventItem in events.Items)
                {
                    DateTime EventDate = new DateTime();
                    TimeSpan? StartTime = null;
                    TimeSpan? EndTime = null;
                    DateTime CreatedDate = new DateTime();

                    string Phone = string.Empty, FamilyName = string.Empty, Category = string.Empty, Address = string.Empty;

                    var EventNameList = eventItem.Summary?.Split('-').ToList();
                    var LocationList = eventItem.Location?.Split('-').ToList();

                    if (eventItem.Start.DateTime != null)
                    {
                        EventDate = eventItem.Start.DateTime.Value.Date;
                        StartTime = eventItem.Start.DateTime.Value.TimeOfDay;
                    }

                    if (eventItem.End.DateTime != null)
                    {
                        EndTime = eventItem.End.DateTime.Value.TimeOfDay;
                    }

                    if (eventItem.Created.Value != null)
                    {
                        CreatedDate = eventItem.Created.Value;

                    }

                    if (eventItem.Description != null)
                    {
                        Phone = eventItem.Description.Replace("Phone:", "").Replace("\n", "").Replace("\r", "").Trim();

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
                    var ETag = Regex.Replace(eventItem.ETag, @"[\W_]", "");

                    CalendarEvent ce = new CalendarEvent
                    {
                        //CalendarName = events.Summary,              
                        EventID = eventItem.Id,
                        EventDate = EventDate,
                        StartTime = StartTime,
                        EndTime = EndTime,
                        FamilyName = FamilyName,
                        Location = eventItem.Location,
                        Address = Address,
                        Phone = Phone,
                        Category = Category,
                        CreatedDate = CreatedDate,
                        ETagID = ETag
                    };
                    CalendarEvents.Add(ce);
                }
            }
            return CalendarEvents;
        }
    }

}
