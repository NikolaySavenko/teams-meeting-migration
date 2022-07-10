using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
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
            var retryOptions = new RetryOptions(TimeSpan.FromSeconds(5), 3);
            await context.CallActivityWithRetryAsync(nameof(EmailSender.EmailSender.SendUpcomingMigrationEmail), retryOptions, user);
            if (!context.IsReplaying) log.LogInformation("[Migration] Started mailbox orchestration for user {UserUserPrincipalName}", user.UserPrincipalName);
            await context.CallSubOrchestratorAsync(nameof(EventsOrchestration.RunEventsOrchestration), user);
            await context.CallActivityWithRetryAsync(nameof(EmailSender.EmailSender.SendMigrationDoneEmail), retryOptions, user);
            // Here can be described another orchestrations...
        }
    }
}
