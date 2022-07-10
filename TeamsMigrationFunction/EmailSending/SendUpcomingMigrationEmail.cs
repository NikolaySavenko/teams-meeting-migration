using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using SendGrid.Helpers.Mail;
using Services;
using System.Threading.Tasks;
using EmailAddress = SendGrid.Helpers.Mail.EmailAddress;


namespace TeamsMigrationFunction.EmailSender
{
    public static partial class EmailSender
    {
        [FunctionName(nameof(SendUpcomingMigrationEmail))]
        public static async Task SendUpcomingMigrationEmail(
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
            await EmailService.SendEmailAsync(user.UserPrincipalName, subject, body);
            log.LogInformation("[Migration] Successfully sent upcoming migration email to {UserUserPrincipalName}", user.UserPrincipalName);
        }
    }
}
