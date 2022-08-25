using System.Globalization;
using Microsoft.Graph;

namespace Services
{
    public class TenantGraphClient
    {
        private readonly GraphServiceClient _graphClient;

        public TenantGraphClient(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<Event> RecreateOnlineMeetingEvent(Event remappedEvent)
        {
            return await _graphClient
                .Users[remappedEvent.Organizer.EmailAddress.Address]
                .Events
                .Request()
                .AddAsync(new Event
                {
                    AllowNewTimeProposals = remappedEvent.AllowNewTimeProposals,
                    Attendees = remappedEvent.Attendees,
                    End = remappedEvent.End,
                    Importance = remappedEvent.Importance,
                    IsAllDay = remappedEvent.IsAllDay,
                    IsCancelled = remappedEvent.IsCancelled,
                    IsDraft = remappedEvent.IsDraft,
                    IsOnlineMeeting = true,
                    IsOrganizer = remappedEvent.IsOrganizer,
                    IsReminderOn = remappedEvent.IsReminderOn,
                    Location = remappedEvent.Location,
                    Locations = remappedEvent.Locations,
                    OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness,
                    OriginalEndTimeZone = remappedEvent.OriginalEndTimeZone,
                    OriginalStart = remappedEvent.OriginalStart,
                    OriginalStartTimeZone = remappedEvent.OriginalStartTimeZone,
                    Recurrence = remappedEvent.Recurrence,
                    ReminderMinutesBeforeStart = remappedEvent.ReminderMinutesBeforeStart,
                    ResponseRequested = remappedEvent.ResponseRequested,
                    Sensitivity = remappedEvent.Sensitivity,
                    ShowAs = remappedEvent.ShowAs,
                    Start = remappedEvent.Start,
                    Body = new ItemBody
                    {
                        Content = " ",
                        ContentType = BodyType.Text
                    },
                    Subject = remappedEvent.Subject,
                    Type = remappedEvent.Type,
                    Categories = remappedEvent.Categories
                });
        }

        public async Task CancelDeprecatedEvent(Event @event)
        {
            const string comment = "Canceled by daemon";
            await _graphClient
                .Users[@event.Organizer.EmailAddress.Address]
                .Events[@event.Id]
                .Cancel(comment)
                .Request()
                .PostAsync();
        }
        
        
        public async Task<IEnumerable<User>> GetAllUsers()
        {
            var users = new List<User>();
            var pagedUsers = await _graphClient.Users
                .Request()
                .GetAsync();
            var pageIterator = PageIterator<User>
                .CreatePageIterator(
                    _graphClient,
                    pagedUsers,
                    // Callback executed for each item in
                    // the collection
                    u =>
                    {
                        users.Add(u);
                        return true;
                    }
                );
            await pageIterator.IterateAsync();
            return users;
        }

        public async Task<IEnumerable<Event>> GetMeetingEventsByUser(User user, string dateTimeFrom)
        {
            var events = new List<Event>();
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("filter", $"isOrganizer eq true and start/dateTime gt '{dateTimeFrom}'"),
                new QueryOption("$count", "true")
            };

            var pagedEvents = await _graphClient.Users[user.Id].Events
                .Request( queryOptions )
                .Filter($"isOrganizer eq true and start/dateTime gt '{dateTimeFrom}'")
                .Top(999)
                .GetAsync();
            
            var pageIterator = PageIterator<Event>
                .CreatePageIterator(
                    _graphClient,
                    pagedEvents,
                    // Callback executed for each item in
                    // the collection
                    e =>
                    {
                        var startTime = DateTime.Parse(e.Start.DateTime);
                        var timeFrom = DateTime.Parse(dateTimeFrom);
                        if ((e.IsOrganizer ?? false) && (e.IsOnlineMeeting ?? false) && (!e.IsCancelled ?? true ) && startTime > timeFrom) { 
                            events.Add(e);
                        }
                        return true;
                    }
                );
            await pageIterator.IterateAsync();
            return events;
        }

        public async Task<Event> UpdateEvent(Event @event)
        {
            return await _graphClient.Users[@event.Organizer.EmailAddress.Address]
                .Events[@event.Id]
                .Request()
                .UpdateAsync(new Event
                {
                    Body = @event.Body
                });
        }
        
        public async Task SendEmail(string userUserPrincipalName, string subject, string body)
        {
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = userUserPrincipalName
                        }
                    }
                }
            };
            await _graphClient.Users["admin@7jhrzx.onmicrosoft.com"]
                .SendMail(message)
                .Request()
                .PostAsync();
        }

        public async Task<User> GetUserByUpn(string upn)
        {
            return await _graphClient.Users[upn]
                .Request()
                .GetAsync();
        }

        public async Task<int> GetUserMeetingsQty(User user, string dateTimeFrom)
        {
            var meetings = await GetMeetingEventsByUser(user, dateTimeFrom);
            return meetings.Count();
        }
    }
}
