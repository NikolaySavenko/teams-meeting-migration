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
                new("startdatetime", dateTimeFrom),
                new("enddatetime", DateTime.MaxValue.ToString())
            };
            var pagedEvents = await _graphClient.Users[user.Id].Events
                .Request(queryOptions)
                .GetAsync();
            var pageIterator = PageIterator<Event>
                .CreatePageIterator(
                    _graphClient,
                    pagedEvents,
                    // Callback executed for each item in
                    // the collection
                    e =>
                    {
                        if ((e.IsOrganizer ?? false) && (e.IsOnlineMeeting ?? false) && (!e.IsCancelled ?? true )) { 
                            events.Add(e);
                        }
                        return true;
                    }
                );
            await pageIterator.IterateAsync();
            return events;
        }
    }
}
