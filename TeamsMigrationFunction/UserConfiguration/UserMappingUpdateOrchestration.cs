using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;

namespace TeamsMigrationFunction.UserConfiguration
{
    public static class UserMappingOrchestration
    {
        [FunctionName(nameof(StartUserMappingUpdate))]
        public static async Task StartUserMappingUpdate(
            [BlobTrigger("meeting-migration-service/user-mapping.csv", Connection = "AzureWebJobsStorage")] Stream csvStream,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            var reader = new StreamReader(csvStream);
            var csv = await reader.ReadToEndAsync();
            log.LogInformation($"[Migration] Trying to launch user mapping update");
            var instanceId = await starter.StartNewAsync(nameof(RunUsersMappingUpdateOrchestration), null, csv);
            log.LogInformation("[Migration] Started user mapping update orchestration with with ID = \'{InstanceId}\'", instanceId);
        }
        
        [FunctionName(nameof(RunUsersMappingUpdateOrchestration))]
        public static async Task RunUsersMappingUpdateOrchestration(
            [OrchestrationTrigger] IDurableOrchestrationContext context,
            ILogger log)
        {
            var input = context.GetInput<string>();
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Started user mapping update orchestration");
            
            var mappingsDictionary = ReadMappingFromCsv(input);
            var mappings = mappingsDictionary.Select(
                mapping => new UserMapping(mapping.Key, mapping.Value)
                );
            await context.CallActivityAsync(nameof(UpdateUsersMapping), mappings);
            // Other orchestrations
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Finished user mapping update orchestration");
        }

        [FunctionName(nameof(UpdateUsersMapping))]
        public static async Task UpdateUsersMapping(
            [ActivityTrigger] UserMapping[] mappingsToUpdate,
            [CosmosDB(
                databaseName: "MeetingMigrationService",
                containerName: "UserMappings",
                Connection = "CosmosDBConnection",
                PartitionKey = "/id",
                CreateIfNotExists = true)] IAsyncCollector<UserMapping> mappings,
            ILogger log)
        {
            log.LogInformation("[Migration] Trying to update users mappings");
            foreach (var mappingToUpdate in mappingsToUpdate)
            {
                log.LogInformation("[Migration] Updating mapping {MappingToUpdate}", mappingToUpdate);
                await mappings.AddAsync(mappingToUpdate);
            }
            log.LogInformation("[Migration] Updated users mappings");
        }

        private static IDictionary<string, string> ReadMappingFromCsv(string csv)
        {
            return new Dictionary<string, string>(
                csv.Split(Environment.NewLine)
                    .Select(line => {
                        var upns = line.Split(",");
                        return KeyValuePair.Create(upns[0], upns[1]);
                    })
            );
        }
    }
}
