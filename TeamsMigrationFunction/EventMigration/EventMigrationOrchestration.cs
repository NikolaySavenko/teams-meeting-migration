using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using TeamsMigrationFunction.UserConfiguration;

namespace TeamsMigrationFunction.EventMigration
{
    public class EventMigrationOrchestration
    {
        private readonly TenantGraphClient _tenantClient;

        public EventMigrationOrchestration(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }

        [FunctionName(nameof(RunEventMigrationOrchestration))]
        public async Task<Event> RunEventMigrationOrchestration(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log)
        {
            var sourceEvent = context.GetInput<Event>();
            if (!context.IsReplaying) log.LogInformation("[Migration] Migrating event with subject: {SourceEventSubject}", sourceEvent.Subject);

            var organizedMapping = await context.CallActivityAsync<UserMapping?>(nameof(GetMappingForSourceUpn), sourceEvent.Organizer.EmailAddress.Address);
            
            if (organizedMapping != null)
            {
                sourceEvent.Organizer.EmailAddress.Address = organizedMapping.DestinationUpn;
                if (!context.IsReplaying) log.LogInformation("[Migration] Meeting: {Subject} -> Found organizer mapping: {Mapping}", sourceEvent.Subject, organizedMapping);
            }
            else
            {
                if (!context.IsReplaying) log.LogInformation("[Migration] Meeting: {Subject} -> Not found organizer mapping: {Email}", sourceEvent.Subject, sourceEvent.Organizer.EmailAddress.Address);
            }
            
            var attendeeMappings = await Task.WhenAll(
                sourceEvent.Attendees
                    .Select(attendee => context.CallActivityAsync<UserMapping>(nameof(GetMappingForSourceUpn), attendee.EmailAddress.Address))
            );

            sourceEvent.Attendees = sourceEvent.Attendees.Select(attendee =>
            {
                var mapping = attendeeMappings.FirstOrDefault(mapping => mapping?.SourceUpn == attendee.EmailAddress.Address);
                if (mapping != null)
                {
                    attendee.EmailAddress.Address = mapping.DestinationUpn;
                    if (!context.IsReplaying) log.LogInformation("[Migration] Meeting: {Subject} -> Found attendee mapping: {Mapping}", sourceEvent.Subject , mapping);
                }
                else
                {
                    if (!context.IsReplaying) log.LogInformation("[Migration] Meeting: {Subject} -> Not found attendee mapping for {Email}", sourceEvent.Subject, attendee.EmailAddress.Address);
                }
                return attendee;
            });
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Finished migrating event with subject: {SourceEventSubject}", sourceEvent.Subject);

            var recreatedEvent = await context.CallActivityAsync<Event>(nameof(RecreateOnlineMeetingEvent), sourceEvent);
            //var updatedEvent = await context.CallActivityAsync<Event>(nameof(UpdateMeetingBody), recreatedEvent);
            //await context.CallActivityAsync(nameof(CancelDeprecatedEvent), sourceEvent);
            return recreatedEvent;
        }

        [FunctionName(nameof(GetMappingForSourceUpn))]
        public static UserMapping GetMappingForSourceUpn(
            [ActivityTrigger] string upn,
            [CosmosDB(
                databaseName: "MeetingMigrationService",
                containerName: "UserMappings",
                Connection = "CosmosDBConnection",
                SqlQuery = "SELECT * FROM c WHERE c.SourceUpn = {upn}"
                )] IEnumerable<UserMapping> userMappings)
        {
            return userMappings.FirstOrDefault();
        }

        [FunctionName(nameof(RecreateOnlineMeetingEvent))]
        public async Task<Event> RecreateOnlineMeetingEvent([ActivityTrigger] Event sourceEvent)
        {
            return await _tenantClient.RecreateOnlineMeetingEvent(sourceEvent);
        }
        
        [FunctionName(nameof(UpdateMeetingBody))]
        public async Task<Event> UpdateMeetingBody([ActivityTrigger] Event @event)
        {
            var optionsUrl = GetOptionsUrlFromOldBody(@event.Body.Content);
            @event.Body.Content = $@"<html>
    <head>
        <meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">
    </head>
    <body>
        <div>
            <br>
            <br>
            <br>
            <div style=""width:100%; height:20px"">
                <span style=""white-space:nowrap; color:#5F5F5F; opacity:.36"">________________________________________________________________________________</span>
            </div>
            <div class=""me-email-text"" lang=""en-US"" style=""color:#252424; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif"">
                <div style=""margin-top:24px; margin-bottom:20px"">
                    <span style=""font-size:24px; color:#252424"">Microsoft Teams meeting</span>
                </div>
                <div style=""margin-bottom:20px"">
                    <div style=""margin-top:0px; margin-bottom:0px; font-weight:bold"">
                        <span style=""font-size:14px; color:#252424"">Join on your computer or mobile app</span>
                    </div>
                    <a href=""{@event.OnlineMeeting.JoinUrl}"">Click here to join the meeting</a>
                </div>
                <div style=""margin-bottom:24px; margin-top:20px"">
                    <a href=""https://aka.ms/JoinTeamsMeeting"" class=""me-email-link"" style=""font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif"">Learn More</a> | <a href=""{optionsUrl}"" class=""me-email-link"" style=""font-size:14px; text-decoration:underline; color:#6264a7; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif"">Meeting options</a>
                 </div></div><div style=""font-size:14px; margin-bottom:4px; font-family:'Segoe UI','Helvetica Neue',Helvetica,Arial,sans-serif"">
                </div>
                <div style=""font-size:12px"">
                </div>
            </div>
            <div style=""width:100%; height:20px"">
                <span style=""white-space:nowrap; color:#5F5F5F; opacity:.36"">________________________________________________________________________________</span>
            </div>
                <div>

                </div>
            </body>
                </html>";
            return await _tenantClient.UpdateEvent(@event);
        }

        private string GetOptionsUrlFromOldBody(string bodyContent)
        {
            var startIndex = bodyContent.IndexOf("https://teams.microsoft.com/meetingOptions/?organizerId");
            var endIndex = bodyContent.IndexOf("language=");
            return bodyContent.Substring(startIndex, endIndex - startIndex + 14);
        }

        [FunctionName(nameof(CancelDeprecatedEvent))]
        public async Task CancelDeprecatedEvent([ActivityTrigger] Event @event, ILogger log)
        {
            try
            {
                await _tenantClient.CancelDeprecatedEvent(@event);
            }
            catch (ServiceException e)
            {
                log.LogError("[Migration] Failed to cancel meeting with subject {EventSubject} for user {UserName} with id {EventId}", @event.Subject, @event.Organizer.EmailAddress.Address, @event.Id);
            }
        }
    }
}
