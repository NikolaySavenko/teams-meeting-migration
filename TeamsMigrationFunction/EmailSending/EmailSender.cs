using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using System.Threading.Tasks;

namespace TeamsMigrationFunction.EmailSending
{
    public class EmailSender
    {
        private readonly TenantGraphClient _tenantClient;

        public EmailSender(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }
        
        [FunctionName(nameof(SendMigrationDoneEmail))]
        public async Task SendMigrationDoneEmail(
            [ActivityTrigger] User user,
            ILogger log)
        {
            log.LogInformation("[Migration] Successfully sent finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
            const string subject = "Your finished migration in Teams";
            const string body = @"Your Teams meetings has finished.";
            await _tenantClient.SendEmail(user.UserPrincipalName, subject, body);
            log.LogInformation("[Migration] Successfully sent finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
        }
        
        [FunctionName(nameof(SendUpcomingMigrationEmail))]
        public async Task SendUpcomingMigrationEmail(
            [ActivityTrigger] User user,
            ILogger log)
        {
            log.LogInformation("[Migration] Trying to send upcoming migration email to {UserUserPrincipalName}", user.UserPrincipalName);
            const string subject = "Your upcoming migration to Teams";
            const string body = @"Your account is under migration to Teams.
App will do this steps:
1. Recreate every meeting in Teams with updated user names.
2. Cancel old meetings.
You will be notified when your migration is done.";
            await _tenantClient.SendEmail(user.UserPrincipalName, subject, body);
            log.LogInformation("[Migration] Successfully sent upcoming migration email to {UserUserPrincipalName}", user.UserPrincipalName);
        }
    }
}
