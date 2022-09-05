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
            var userConfigurations = await context.CallActivityAsync<Dictionary<string, string>>(nameof(ReadUserConfigurationsFromCsv), csv);
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
            
            foreach (var user in users)
            {
                try
                {
                    await context.CallSubOrchestratorAsync(MailboxMigrationOrchestratorName, user);
                }
                catch (Exception e)
                {
                    if (!context.IsReplaying) log.LogError($"Failed to migrate mailbox for {user} with exception {e}");
                }
            }
            
            // await Task.WhenAll(
            //     users.Select(
            //         upn => context.CallSubOrchestratorAsync(MailboxMigrationOrchestratorName, upn)
            //     )
            // );
        }

        [FunctionName(nameof(GetUsersForConfigurations))]
        public async Task<IEnumerable<User>> GetUsersForConfigurations([ActivityTrigger] IDurableActivityContext context)
        {
            return await _tenantClient.GetAllUsers();
        }
        
        [FunctionName(nameof(ReadUserConfigurationsFromCsv))]
        public static Task<Dictionary<string, string>> ReadUserConfigurationsFromCsv([ActivityTrigger] string csv)
        {
            var lines = csv.Split(Environment.NewLine);
            // Assert missing key or value
            var incorrectLinesSb = new StringBuilder();
            foreach (var line in lines)
            {
                var upn2DateTime = line.Split(",");
                if (upn2DateTime.Length != 2 || string.IsNullOrEmpty(upn2DateTime[0]) || string.IsNullOrEmpty(upn2DateTime[1]))
                {
                    var index = Array.IndexOf(lines, line);
                    incorrectLinesSb.AppendLine($"[{index}: ({line})]");
                }
            }
            
            if (incorrectLinesSb.Length > 0)
            {
                throw new InvalidDataException($"Invalid input configuration CSV: \n {incorrectLinesSb}");
            }
            
            return Task.FromResult(new Dictionary<string, string>(
                csv.Split(Environment.NewLine)
                    .Select(line =>
                    {
                        var parsedParams = line.Split(",");
                        return KeyValuePair.Create(parsedParams[0], parsedParams[1]);
                    }))
            );
        }
    }
}
