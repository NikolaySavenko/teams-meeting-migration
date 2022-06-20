using System;
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
    public class EventsOrchestration
    {
        private readonly TenantGraphClient _tenantClient;

        public EventsOrchestration(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }
        
        [FunctionName(nameof(RunEventsOrchestration))]
        public async Task<IEnumerable<Event>> RunEventsOrchestration(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log)
        {
            var user = context.GetInput<User>();
            
            await ConfigureUserMailboxStartTime(context, user);

            var organizedEvents = await context.CallActivityAsync<Event[]>(nameof(GetMeetingEventsOrganizedByUser), user);
            if (!context.IsReplaying) log.LogInformation("[Migration] Found {OrganizedEventsLength} events for user: {UserUserPrincipalName}. Starting migration", organizedEvents.Length, user.UserPrincipalName);
            return await Task.WhenAll(
                organizedEvents.Select(
                    @event => context.CallSubOrchestratorAsync<Event>(nameof(EventMigrationOrchestration.RunEventMigrationOrchestration), @event)
                )
            );
        }

        private static async Task ConfigureUserMailboxStartTime(IDurableOrchestrationContext context, User user)
        {
            var entityId = new EntityId(nameof(UserConfiguration.UserConfiguration), user.UserPrincipalName);
            var configProxy = context.CreateEntityProxy<IUserConfiguration>(entityId);
            var mailboxStartTime = await configProxy.GetMailboxStartTime();
            user.AdditionalData ??= new Dictionary<string, object>();
            user.AdditionalData.Add(UserConfiguration.UserConfiguration.UserAdditionalFieldName, mailboxStartTime);
        }

        [FunctionName(nameof(GetMeetingEventsOrganizedByUser))]
        public async Task<IEnumerable<Event>> GetMeetingEventsOrganizedByUser([ActivityTrigger] User user, ILogger log)
        {
            try
            {
                var dateTimeFrom = user.AdditionalData[UserConfiguration.UserConfiguration.UserAdditionalFieldName] as string;
                log.LogInformation("[Migration] Scanning meeting for {UserName} from {StartTime}", user.UserPrincipalName, dateTimeFrom);
                return await _tenantClient.GetMeetingEventsByUser(user, dateTimeFrom);
            }
            catch (ServiceException e)
            {
                log.LogError(e, "[Migration] Cannot get events for user {UserUserPrincipalName}", user.UserPrincipalName);
            }
            return Array.Empty<Event>();
        }
    }
}
