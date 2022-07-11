using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using TeamsMigrationFunction.EmailSending;
using TeamsMigrationFunction.EventMigration;

namespace TeamsMigrationFunction.MailboxOrchestration
{
    public static class MailboxOrchestration
    {
        [FunctionName(nameof(RunMailboxOrchestrator))]
        public static async Task RunMailboxOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log
        )
        {
            var user = context.GetInput<User>();
            var retryOptions = new RetryOptions(TimeSpan.FromSeconds(5), 1);
            try
            {
                await context.CallActivityWithRetryAsync(nameof(EmailSender.SendUpcomingMigrationEmail), retryOptions, user);
            }
            catch (Exception ex)
            {
                log.LogError("[Migration] Failed to sent upcoming migration email to {userPrincipalName}", user.UserPrincipalName);
            }
            if (!context.IsReplaying) log.LogInformation("[Migration] Started mailbox orchestration for user {UserUserPrincipalName}", user.UserPrincipalName);
            await context.CallSubOrchestratorAsync(nameof(EventsOrchestration.RunEventsOrchestration), user);
            try
            {
                await context.CallActivityWithRetryAsync(nameof(EmailSender.SendMigrationDoneEmail), retryOptions, user);
            }
            catch (Exception ex)
            {
                log.LogError("[Migration] Failed to sent migration done email to {userPrincipalName}", user.UserPrincipalName);
            }
            
            // Here can be described another orchestrations...
        }
    }
}
