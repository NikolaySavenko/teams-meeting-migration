using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using System.Threading.Tasks;

namespace TeamsMigrationFunction.EmailSender
{
    public static partial class EmailSender
    {
        [FunctionName(nameof(SendMigrationDoneEmail))]
        public static async Task SendMigrationDoneEmail(
            [ActivityTrigger] User user,
            ILogger log)
        {
            log.LogInformation("[Migration] Successfully sent finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
            const string subject = "Your finished migration in Teams";
            const string body = @"Your Teams meetings has finished.";
            await EmailService.SendEmailAsync(user.UserPrincipalName, subject, body);
            log.LogInformation("[Migration] Successfully sent finished migration email to {UserUserPrincipalName}", user.UserPrincipalName);
        }
    }
}
