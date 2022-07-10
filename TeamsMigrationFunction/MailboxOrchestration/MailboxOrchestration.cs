using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
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
            await context.CallActivityAsync(nameof(EmailSender.EmailSender.SendUpcomingMigrationEmail), user);
            if (!context.IsReplaying) log.LogInformation("[Migration] Started mailbox orchestration for user {UserUserPrincipalName}", user.UserPrincipalName);
            await context.CallSubOrchestratorAsync(nameof(EventsOrchestration.RunEventsOrchestration), user);
            await context.CallActivityAsync(nameof(EmailSender.EmailSender.SendMigrationDoneEmail), user);
            // Here can be described another orchestrations...
        }
    }
}
