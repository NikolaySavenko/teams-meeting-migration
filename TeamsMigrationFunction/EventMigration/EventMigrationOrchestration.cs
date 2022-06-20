using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using TeamsMigrationFunction.UserMapping;

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
            var mapperEntityId = new EntityId(nameof(UserMapper), "global");
            var mapperProxy = context.CreateEntityProxy<IUserMapper>(mapperEntityId);
            
            sourceEvent.Organizer.EmailAddress.Address = await mapperProxy.GetUserDestinationUpn(sourceEvent.Organizer.EmailAddress.Address);
            sourceEvent.Attendees = await Task.WhenAll(
                sourceEvent.Attendees
                    .Select(attendee => mapperProxy.GetDestinationAttendee(attendee))
            );
            
            var recreatedEvent = await context.CallActivityAsync<Event>(nameof(RecreateOnlineMeetingEvent), sourceEvent);
            await context.CallActivityAsync(nameof(CancelDeprecatedEvent), sourceEvent);
            return recreatedEvent;
        }

        [FunctionName(nameof(RecreateOnlineMeetingEvent))]
        public async Task<Event> RecreateOnlineMeetingEvent([ActivityTrigger] Event sourceEvent)
        {
            return await _tenantClient.RecreateOnlineMeetingEvent(sourceEvent);
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
