using System;
using System.Collections.Generic;
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
using KeyValuePair = System.Collections.Generic.KeyValuePair;

namespace TeamsMigrationFunction.UsersOrchestration
{
    public class UsersMigrationOrchestration
    {
        private readonly TenantGraphClient _tenantClient;

        private const string MailboxMigrationOrchestratorName = nameof(MailboxOrchestration.MailboxOrchestration.RunMailboxOrchestrator);
        
        public UsersMigrationOrchestration(TenantGraphClient tenantClient)
        {
            _tenantClient = tenantClient;
        }
        
        [FunctionName(nameof(RunUsersMigrationOrchestrationHttp))]
        public static async Task<HttpResponseMessage> RunUsersMigrationOrchestrationHttp(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestMessage req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log
        )
        {
            var csv = await req.Content.ReadAsStringAsync();
            if (string.IsNullOrEmpty(csv))
            {
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("No such file", Encoding.UTF8, "application/text")
                };
            }
            var instanceId = await starter.StartNewAsync(nameof(RunUsersMigrationOrchestration), null, csv);
            log.LogInformation("[Migration] Started migration with ID = \'{InstanceId}\'", instanceId);
            return starter.CreateCheckStatusResponse(req, instanceId);
        }

        [FunctionName(nameof(RunUsersMigrationOrchestration))]
        public async Task RunUsersMigrationOrchestration(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log)
        {
            var csv = context.GetInput<string>();
            var userConfigurations = ReadUserConfigurationsFromCsv(csv);
            var users = userConfigurations.Keys;
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Found {MailboxConfigs} user configs, preparing for orchestration...", userConfigurations.Count);
            
            await Task.WhenAll(
                users.Select(
                    upn => {
                        var configEntityId = new EntityId(nameof(UserConfiguration), upn);
                        var configProxy = context.CreateEntityProxy<IUserConfiguration>(configEntityId);
                        var hasTime = userConfigurations.TryGetValue(upn, out var startTime);
                        
                        return configProxy.SetMailboxStartTime(hasTime ? startTime : DateTime.MinValue.ToString());
                    })
                );
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Found {UsersLength} users. Starting mailbox orchestration...", users.Count);
            await Task.WhenAll(
                users.Select(
                    upn => context.CallSubOrchestratorAsync(MailboxMigrationOrchestratorName, upn)
                )
            );
        }

        [FunctionName(nameof(GetUsersForConfigurations))]
        public async Task<IEnumerable<User>> GetUsersForConfigurations([ActivityTrigger] IDurableActivityContext context)
        {
            return await _tenantClient.GetAllUsers();
        }
        
        private static IDictionary<string, string> ReadUserConfigurationsFromCsv(string csv)
        {
            return new Dictionary<string, string>(
                csv.Split(Environment.NewLine)
                    .Select(line => {
                        var parsedParams = line.Split(",");
                        return KeyValuePair.Create(parsedParams[0], parsedParams[1]);
                    })
            );
        }
    }
}
