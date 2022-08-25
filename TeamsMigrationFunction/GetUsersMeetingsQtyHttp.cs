using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using TeamsMigrationFunction.UserConfiguration;
using TeamsMigrationFunction.UsersOrchestration;

namespace TeamsMigrationFunction
{
    public class GetUsersMeetingsQtyHttp
    {
        
        private readonly TenantGraphClient _tenantClient;

        public GetUsersMeetingsQtyHttp(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }
        
        [FunctionName("GetUsersMeetingsQtyHttp")]
        public static async Task<HttpResponseMessage> RunAsync([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            var csv = await req.Content.ReadAsStringAsync();
            if (string.IsNullOrEmpty(csv))
            {
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("No such file", Encoding.UTF8, "application/text")
                };
            }
            var instanceId = await starter.StartNewAsync(nameof(RunMeetingsCounterOrhestration), null, csv);
            log.LogInformation("[Migration] Started counting with ID = \'{InstanceId}\'", instanceId);
            return starter.CreateCheckStatusResponse(req, instanceId);
        }

        [FunctionName(nameof(RunMeetingsCounterOrhestration))]
        public async Task RunMeetingsCounterOrhestration(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log)
        {
            var csv = context.GetInput<string>();
            var userConfigurations = await context.CallActivityAsync<Dictionary<string, string>>(nameof(UsersMigrationOrchestration.ReadUserConfigurationsFromCsv), csv);
            var users = userConfigurations.Keys;
            
            await Task.WhenAll(
                users.Select(
                    upn => {
                        var configEntityId = new EntityId(nameof(UserConfiguration), upn);
                        var configProxy = context.CreateEntityProxy<IUserConfiguration>(configEntityId);
                        var hasTime = userConfigurations.TryGetValue(upn, out var startTime);
                        
                        return configProxy.SetMailboxStartTime(hasTime ? startTime : DateTime.MinValue.ToString());
                    })
            );
            
            await Task.WhenAll(
                users.Select(
                    upn => context.CallSubOrchestratorAsync(nameof(PrintUserMeetingsQty), upn)
                )
            );
        }
        
        [FunctionName(nameof(PrintUserMeetingsQty))]
        public async Task PrintUserMeetingsQty(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log
        )
        {
            var normalLogger = context.CreateReplaySafeLogger(log);
            
            var upn = context.GetInput<string>();
            var user = await context.CallActivityAsync<User>(nameof(MailboxOrchestration.MailboxOrchestration.GetUserByUpn), upn);
            var entityId = new EntityId(nameof(UserConfiguration.UserConfiguration), user.UserPrincipalName);
            var configProxy = context.CreateEntityProxy<IUserConfiguration>(entityId);
            var mailboxStartTime = await configProxy.GetMailboxStartTime();

            normalLogger.LogInformation("[Migration] Scanning meeting for {UserName} from {StartTime}", user.UserPrincipalName, mailboxStartTime);
            var count = await context.CallActivityAsync<int>(nameof(GetMeetingQtyForUser), (user, mailboxStartTime));
            normalLogger.LogInformation("[Migration] Found {OrganizedEventsLength} events for user: {UserUserPrincipalName}", count, user.UserPrincipalName);
            // Here can be described another orchestrations...
        }

        [FunctionName(nameof(GetMeetingQtyForUser))]
         public async Task<int> GetMeetingQtyForUser([ActivityTrigger]IDurableActivityContext inputs)
         {
             (User user, string time) studentInfo = inputs.GetInput<(User, string)>();

             return await _tenantClient.GetUserMeetingsQty(studentInfo.user, studentInfo.time);
         }
    }
}
