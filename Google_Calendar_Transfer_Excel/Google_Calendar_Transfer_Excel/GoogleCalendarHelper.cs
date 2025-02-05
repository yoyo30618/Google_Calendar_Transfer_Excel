using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

public class GoogleCalendarHelper
{
    static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
    static string ApplicationName = "Google Calendar Transfer";

    public static async Task<CalendarService> GetCalendarServiceAsync()
    {
        using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        {
            string credPath = "token.json"; // OAuth Token 儲存
            var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.FromStream(stream).Secrets,
                Scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true));

            return new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
    }
    public static async Task<List<Event>> GetCalendarEventsForSpecificCalendar(string calendarId, DateTime startDate, DateTime endDate)
    {
        var service = await GetCalendarServiceAsync();

        EventsResource.ListRequest request = service.Events.List(calendarId);
        request.TimeMin = startDate;
        request.TimeMax = endDate;
        request.ShowDeleted = false;
        request.SingleEvents = true;
        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

        Events events = await request.ExecuteAsync();
        return events.Items != null ? new List<Event>(events.Items) : new List<Event>();
    }

    public static async Task<List<CalendarListEntry>> GetAllCalendarsAsync()
    {
        var service = await GetCalendarServiceAsync();

        // List the calendars associated with the account
        var calendarListRequest = service.CalendarList.List();
        var calendarList = await calendarListRequest.ExecuteAsync();

        return calendarList.Items.ToList(); // 返回日曆列表
    }
    public class CalendarItem
    {
        public string CalendarName { get; set; }
        public string CalendarId { get; set; }

        public CalendarItem(string name, string id)
        {
            CalendarName = name;
            CalendarId = id;
        }

        public override string ToString()
        {
            return CalendarName; // 顯示名稱
        }
    }

}
