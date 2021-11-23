using Microsoft.Graph;
using TimeZoneConverter;

namespace simple_graph_console
{
    public class CalendarHelper
    {
        private static GraphServiceClient graphClient;

        public static void Initialize(GraphServiceClient client)
        {
            graphClient = client;
        }

        public static async Task<Calendar> CreateCalendar(string calendarName)
        {
            try
            {
                var calendar = new Calendar
                {
                    Name = calendarName
                };

                return await graphClient.Me.Calendars
                    .Request()
                    .AddAsync(calendar);

            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating calendar: {ex.Message}");
                return null;
            }
        }

        public static async Task<Event> CreateEvent(string subject, string content, string location)
        {
            try
            {
                var startTime = DateTime.Now.AddHours(1).ToString();
                var endTime = DateTime.Now.AddHours(2).ToString();

                var @event = new Event
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = content
                    },
                    Start = new DateTimeTimeZone
                    {
                        DateTime = startTime,
                        TimeZone = "SE Asia Standard Time"
                    },
                    End = new DateTimeTimeZone
                    {
                        // DateTime = "2017-04-15T14:00:00",
                        DateTime = endTime,
                        TimeZone = "SE Asia Standard Time"
                    },
                    Location = new Location
                    {
                        DisplayName = location
                    },
                    AllowNewTimeProposals = true,
                };

                return await graphClient.Me.Events
                    .Request()
                    .Header("Prefer", "outlook.timezone=\"SE Asia Standard Time\"")
                    .AddAsync(@event);

            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error creating calendar: {ex.Message}");
                return null;
            }
        }
    }


}