using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using Services;
using TeamsMigrationFunction.EmailSending;
using TeamsMigrationFunction.EventMigration;

namespace TeamsMigrationFunction.MailboxOrchestration
{
    public class MailboxOrchestration
    {
        private readonly TenantGraphClient _tenantClient;
        
        public MailboxOrchestration(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }
        
        [FunctionName(nameof(RunMailboxOrchestrator))]
        public static async Task RunMailboxOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log
        )
        {
            var upn = context.GetInput<string>();
            var user = await context.CallActivityAsync<User>(nameof(GetUserByUpn), upn);
            
            await SendUpcomingMigrationEmail(user, context, log);
            await MigrateMailboxForUser(user, context, log);
            await SendFinishedMigrationEmail(user, context, log);
            // Here can be described another orchestrations...
        }

        [FunctionName(nameof(GetUserByUpn))]
        public async Task<User> GetUserByUpn(
            [ActivityTrigger] string upn,
            ILogger log)
        {
            log.LogInformation("[Migration] Trying to get user with upn: {upn}", upn);
            return await _tenantClient.GetUserByUpn(upn);
        }
        
        private static async Task MigrateMailboxForUser(User user, IDurableOrchestrationContext context, ILogger log)
        {
            if (!context.IsReplaying) log.LogInformation("[Migration] Started mailbox orchestration for user {UserUserPrincipalName}", user.UserPrincipalName);
            await context.CallSubOrchestratorAsync(nameof(EventsOrchestration.RunEventsOrchestration), user);
        }

        private static async Task SendUpcomingMigrationEmail(User user, IDurableOrchestrationContext context, ILogger log)
        {
            try
            {
                if (!context.IsReplaying) log.LogInformation("[Migration] Trying to send upcoming migration email to {UserUserPrincipalName}", user.UserPrincipalName);
                var email = new EmailMessage(
                    "Meeting Migration Service",
                    user.UserPrincipalName,
                    "Your upcoming migration to Teams",
                    @"Your account is under migration to Teams.
App will do this steps:
1. Recreate every meeting in Teams with updated user names.
2. Cancel old meetings.
You will be notified when your migration is done.");
                var request = EmailSender.BuildEmailRequest(email);
                await context.CallHttpAsync(request);
                if (!context.IsReplaying) log.LogInformation("[Migration] Successfully sent upcoming migration email to {UserUserPrincipalName}", user.UserPrincipalName);
            }
            catch (Exception ex)
            {
                log.LogError("[Migration] Failed to sent upcoming migration email to {userPrincipalName}", user.UserPrincipalName);
            }
        }
        
        private static async Task SendFinishedMigrationEmail(User user, IDurableOrchestrationContext context, ILogger log)
        {
            try
            {
                if (!context.IsReplaying) log.LogInformation("[Migration] Trying to send finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
                var email = new EmailMessage(
                    "Meeting Migration Service",
                    user.UserPrincipalName,
                    "Your finished migration in Teams",
                    "Your Teams meetings has finished.");
                var request = EmailSender.BuildEmailRequest(email);
                await context.CallHttpAsync(request);
                if (!context.IsReplaying) log.LogInformation("[Migration] Successfully sent finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
            }
            catch (Exception ex)
            {
                log.LogError("[Migration] Failed to sent upcoming migration email to {userPrincipalName}", user.UserPrincipalName);
            }
        }
    }
}
