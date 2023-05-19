
//using Google.Apis.Calendar.v3;
//using Google.Apis.Calendar.v3.Data;
//using Google.Apis.Auth.OAuth2;
//using Google.Apis.Services;
//using Google.Apis.Util.Store;
//using System.Text.RegularExpressions;
//using Application.Interfaces;
//using Persistence;
//using Domain.Models;
//using Application.Helper;

//namespace Application.Repository
//{

//    public class CalendarEventRepository : ICalendarEventRepository
//    {
//        static string[] Scopes = { CalendarService.Scope.Calendar, CalendarService.Scope.CalendarEvents };
//        static string ApplicationName = "Google Calendar API .NET Quickstart";
//        private DataContext _context;
//        public CalendarEventRepository(DataContext context)
//        {
//            _context = context;
//        }

//        public bool CalendarEventExists(string id)
//        {
//            return _context.CalendarEvents.Any(c => c.EventID == id);
//        }

//        public bool CreateCalendarEvent(CalendarEvent calendarEvent)
//        {
//            _context.Add(calendarEvent);
//            return Save();
//        }

//        public bool DeleteCalendarEvent(CalendarEvent calendarEvent)
//        {
//            _context.Remove(calendarEvent);
//            return Save();
//        }


//        public ICollection<CalendarEvent> GetCalendarEvent()
//        {
//            // try
//            // {


//            var googleDrive = new GoogleDrive();
//            var FuncRet = googleDrive.GetDrive();



//            List<CalendarEvent> NewEvents = new List<CalendarEvent>();
//            List<CalendarEvent> AllEvents = GoogleCalendar.GetCalendarEvents();
//            List<CalendarEvent> OldEvents = _context.CalendarEvents.ToList();


//            //var EventIDs = new List<string>();
//            //var EventIDs = OpenXML.GetCellValues(FuncRet.Item1, "Events", "A");
//            //OpenXML.AddEventToExcel(FuncRet.Item1, AllEvents, EventIDs);

//            googleDrive.UploadFileToDrive(FuncRet.Item1, FuncRet.Item2);


//            var AllEventID = AllEvents.Select(x => x.EventID).ToList();
//            var OldEventID = OldEvents.Select(x => x.EventID).ToList();

//            var result = from a in AllEventID where !OldEventID.Contains(a) select a.ToString();

//            foreach (string item in result)
//            {
//                CalendarEvent? NewEvent = null;
//                NewEvent = AllEvents.FirstOrDefault(x => x.EventID == item);
//                if (NewEvent != null) { NewEvents.Add(NewEvent); }
//            }


//            foreach (CalendarEvent EventItem in AllEvents.Where(x => x.EventDate >= DateTime.Now))
//            {
//                CalendarEvent? DataEvent = OldEvents.Where(x => x.EventID == EventItem.EventID).FirstOrDefault();
//                if (DataEvent != null)
//                {
//                    DataEvent.FamilyName = EventItem.FamilyName;
//                    DataEvent.Location = EventItem.Location;
//                    DataEvent.Phone = EventItem.Phone;
//                    DataEvent.Address = EventItem.Address;
//                    DataEvent.StartTime = EventItem.StartTime;
//                    DataEvent.EndTime = EventItem.EndTime;
//                    DataEvent.ETagID = EventItem.ETagID;
//                }
//            }



//            _context.CalendarEvents.AddRange(NewEvents);
//            _context.SaveChanges();

//            return _context.CalendarEvents.ToList();

//            // }
//            // catch (DbUpdateException ex)
//            //    when ((ex.InnerException as SqlException)?.Number == 2627)
//            // {
//            //     return _context.CalendarEvents.ToList();
//            //     // Handle unique key violation
//            // }

//        }

//        public CalendarEvent GetCalendarEvent(string id)
//        {
//            return _context.CalendarEvents.Where(e => e.EventID == id).FirstOrDefault();
//        }

//        //public ICollection<Pokemon> GetPokemonByCategory(int categoryId)
//        //{
//        //    return _context.PokemonCategories.Where(e => e.CategoryId == categoryId).Select(c => c.Pokemon).ToList();
//        //}

//        public bool Save()
//        {
//            var saved = _context.SaveChanges();
//            return saved > 0 ? true : false;
//        }

//        public bool UpdateCalendarEvent(CalendarEvent calendarEvent)
//        {
//            _context.Update(calendarEvent);
//            return Save();
//        }

//        public CalendarService Authorize()
//        {
//            UserCredential credential;

//            using (var stream =
//                new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
//            {
//                // The file token.json stores the user's access and refresh tokens, and is created
//                // automatically when the authorization flow completes for the first time.
//                string credPath = "token2.json";
//                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
//                    GoogleClientSecrets.FromStream(stream).Secrets,
//                    Scopes,
//                    "user",
//                    CancellationToken.None,
//                    new FileDataStore(credPath, true)).Result;
//                Console.WriteLine("Credential file saved to: " + credPath);
//            }

//            // Create Google Calendar API service.
//            var service = new CalendarService(new BaseClientService.Initializer()
//            {
//                HttpClientInitializer = credential,
//                ApplicationName = ApplicationName,
//            });

//            return service;
//        }

//        public List<CalendarEvent> GetCalendarEvents()
//        {
//            var Service = Authorize();
//            var Events = GetEvents(Service);
//            var CalendarEvents = GetEventDetails(Events);

//            return CalendarEvents;
//        }
//        public CalendarEvent GetEventDetail(Event CalEvent)
//        {

//            DateTime EventDate = new DateTime();
//            TimeSpan? StartTime = null;
//            TimeSpan? EndTime = null;
//            DateTime CreatedDate = new DateTime();

//            string Phone = string.Empty, FamilyName = string.Empty, Category = string.Empty, Address = string.Empty;

//            var EventNameList = CalEvent.Summary?.Split('-').ToList();
//            var LocationList = CalEvent.Location?.Split('-').ToList();

//            if (CalEvent.Start.DateTime != null)
//            {
//                EventDate = CalEvent.Start.DateTime.Value.Date;
//                StartTime = CalEvent.Start.DateTime.Value.TimeOfDay;
//            }

//            if (CalEvent.End.DateTime != null)
//            {
//                EndTime = CalEvent.End.DateTime.Value.TimeOfDay;
//            }

//            if (CalEvent.Created.Value != null)
//            {
//                CreatedDate = CalEvent.Created.Value;

//            }

//            if (CalEvent.Description != null)
//            {
//                Phone = CalEvent.Description.Replace("Phone:", "").Replace("\n", "").Replace("\r", "").Trim();

//            }

//            if (EventNameList != null)
//            {
//                if (EventNameList.Count > 0)
//                {
//                    Category = EventNameList[0].Trim();
//                }
//                if (EventNameList.Count > 1)
//                {
//                    //FamilyName = EventNameList[1].Trim();
//                    FamilyName = EventNameList[1].Replace("\n", "").Replace("\r", "").Trim();
//                }
//            }

//            if (LocationList != null)
//            {
//                if (LocationList.Count() > 1)
//                {
//                    // Location= LocationList[0].Trim();
//                    Address = LocationList[1].Trim();
//                }
//            }
//            var ETag = Regex.Replace(CalEvent.ETag, @"[\W_]", "");

//            CalendarEvent calendarEvent = new CalendarEvent
//            {
//                //CalendarName = events.Summary,              
//                EventID = CalEvent.Id,
//                EventDate = EventDate,
//                StartTime = StartTime,
//                EndTime = EndTime,
//                FamilyName = FamilyName,
//                Location = CalEvent.Location,
//                Address = Address,
//                Phone = Phone,
//                Category = Category,
//                CreatedDate = CreatedDate,
//                ETagID = ETag
//            };
//            return calendarEvent;
//        }

//        public List<CalendarEvent> GetEventDetails(Events events)
//        {
//            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();
//            if (events.Items != null && events.Items.Count > 0)
//            {

//                foreach (var eventItem in events.Items)
//                {
//                    DateTime EventDate = new DateTime();
//                    TimeSpan? StartTime = null;
//                    TimeSpan? EndTime = null;
//                    DateTime CreatedDate = new DateTime();

//                    string Phone = string.Empty, FamilyName = string.Empty, Category = string.Empty, Address = string.Empty, Location = string.Empty;

//                    var EventNameList = eventItem.Summary?.Split('-').ToList();
//                    var LocationList = eventItem.Location?.Split('-').ToList();

//                    if (eventItem.Start.DateTime != null)
//                    {
//                        EventDate = eventItem.Start.DateTime.Value.Date;
//                        StartTime = eventItem.Start.DateTime.Value.TimeOfDay;
//                    }

//                    if (eventItem.End.DateTime != null)
//                    {
//                        EndTime = eventItem.End.DateTime.Value.TimeOfDay;
//                    }

//                    if (eventItem.Created.Value != null)
//                    {
//                        CreatedDate = eventItem.Created.Value;

//                    }

//                    if (eventItem.Description != null)
//                    {
//                        Phone = eventItem.Description.Replace("Phone:", "").Replace("\n", "").Replace("\r", "").Trim();

//                    }

//                    if (EventNameList != null)
//                    {
//                        if (EventNameList.Count > 0)
//                        {
//                            Category = EventNameList[0].Trim();
//                        }
//                        if (EventNameList.Count > 1)
//                        {
//                            //FamilyName = EventNameList[1].Trim();
//                            FamilyName = EventNameList[1].Replace("\n", "").Replace("\r", "").Trim();
//                        }
//                    }

//                    if (LocationList != null)
//                    {
//                        if (LocationList.Count() > 1)
//                        {
//                            Location = LocationList[0].Trim();
//                            Address = LocationList[1].Trim();
//                        }
//                    }
//                    var ETag = Regex.Replace(eventItem.ETag, @"[\W_]", "");

//                    CalendarEvent ce = new CalendarEvent
//                    {
//                        //CalendarName = events.Summary,              
//                        EventID = eventItem.Id,
//                        EventDate = EventDate,
//                        StartTime = StartTime,
//                        EndTime = EndTime,
//                        FamilyName = FamilyName,
//                        Location = Location,
//                        Address = Address,
//                        Phone = Phone,
//                        Category = Category,
//                        CreatedDate = CreatedDate,
//                        ETagID = ETag
//                    };
//                    CalendarEvents.Add(ce);
//                }
//            }
//            return CalendarEvents;
//        }

//        public Events GetEvents(CalendarService service)
//        {
//            List<CalendarEvent> CalendarEvents = new List<CalendarEvent>();

//            EventsResource.ListRequest request = service.Events.List("3dc3debda4eb8d023e70ae94e7043de683c6af12080c8a7515bd4e2101625b4c@group.calendar.google.com");
//            request.TimeMin = DateTime.Parse("1/1/2020"); ;
//            request.TimeMax = DateTime.Parse("12/31/2023"); ;
//            request.ShowDeleted = false;
//            request.SingleEvents = true;
//            //request.MaxResults = 10;
//            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
//            Events events = request.Execute();

//            return events;
//        }



//        public List<CalendarEvent> UpdateEvents(List<CalendarEvent> AllEvents, List<CalendarEvent> OldEvents)
//        {
//            List<CalendarEvent> NewEvents = new List<CalendarEvent>();
//            var AllEventID = AllEvents.Select(x => x.EventID).ToList();
//            var OldEventID = OldEvents.Select(x => x.EventID).ToList();

//            var result = from a in AllEventID where !OldEventID.Contains(a) select a.ToString();

//            foreach (string item in result)
//            {
//                CalendarEvent? NewEvent = null;
//                NewEvent = AllEvents.FirstOrDefault(x => x.EventID == item);
//                if (NewEvent != null) { NewEvents.Add(NewEvent); }
//            }


//            foreach (CalendarEvent EventItem in AllEvents.Where(x => x.EventDate >= DateTime.Now))
//            {
//                CalendarEvent? DataEvent = OldEvents.Where(x => x.EventID == EventItem.EventID).FirstOrDefault();
//                if (DataEvent != null)
//                {
//                    DataEvent.FamilyName = EventItem.FamilyName;
//                    DataEvent.Location = EventItem.Location;
//                    DataEvent.Address = EventItem.Address;
//                    DataEvent.Phone = EventItem.Phone;
//                    DataEvent.StartTime = EventItem.StartTime;
//                    DataEvent.EndTime = EventItem.EndTime;
//                    DataEvent.ETagID = EventItem.ETagID;
//                }
//            }


//            return NewEvents;
//        }


//        public Event CreateEvent(CalendarEvent calendarEvent)
//        {

//            DateTime StartTime = DateTime.Parse(calendarEvent.EventDate.ToString());
//            StartTime = StartTime.AddMinutes(calendarEvent.StartTime.Value.TotalMinutes);

//            DateTime EndTime = DateTime.Parse(calendarEvent.EventDate.ToString());
//            EndTime = EndTime.AddMinutes(calendarEvent.EndTime.Value.TotalMinutes);

//            Event newEvent = new Event()
//            {
//                Description = $"Phone - {calendarEvent.Phone}",
//                Summary = $"{calendarEvent?.Category} - {calendarEvent?.FamilyName}",
//                Location = $"{calendarEvent?.Location} - {calendarEvent?.Address}",

//                Start = new EventDateTime() { DateTime = StartTime, TimeZone = "America/New_York" },
//                End = new EventDateTime() { DateTime = EndTime, TimeZone = "America/New_York" }
//            };
//            return newEvent;
//        }
//        public Event UpdateEvent(Event GoogleEvent, CalendarEvent calendarEvent)
//        {
//            GoogleEvent.Description = $"Phone - {calendarEvent.Phone}";
//            GoogleEvent.Summary = $"{calendarEvent?.Category} - {calendarEvent?.FamilyName}";
//            GoogleEvent.Location = $"{calendarEvent?.Location} - {calendarEvent?.Address}";

//            DateTime StartTime = DateTime.Parse(calendarEvent.EventDate.ToString());
//            StartTime = StartTime.AddMinutes(calendarEvent.StartTime.Value.TotalMinutes);

//            DateTime EndTime = DateTime.Parse(calendarEvent.EventDate.ToString());
//            EndTime = EndTime.AddMinutes(calendarEvent.EndTime.Value.TotalMinutes);

//            GoogleEvent.Start = new EventDateTime() { DateTime = StartTime, TimeZone = "America/New_York" };
//            GoogleEvent.End = new EventDateTime() { DateTime = EndTime, TimeZone = "America/New_York" };

//            return GoogleEvent;
//        }

//    }

//}

















