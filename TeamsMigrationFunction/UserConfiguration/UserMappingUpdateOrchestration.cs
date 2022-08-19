using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.Cosmos;
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
            
            await context.CallActivityAsync(nameof(DeleteOldMappingsContainer), null);
            var mappingsDictionary = await context.CallActivityAsync<Dictionary<string, string>>(nameof(ReadMappingFromCsv), input);
            var mappings = mappingsDictionary.Select(
                mapping => new UserMapping(mapping.Key, mapping.Value)
                );
            await context.CallActivityAsync(nameof(UpdateUsersMapping), mappings);
            // Other orchestrations
            
            if (!context.IsReplaying) log.LogInformation("[Migration] Finished user mapping update orchestration");
        }
        
        [FunctionName(nameof(DeleteOldMappingsContainer))]
        public static async Task DeleteOldMappingsContainer(
            [ActivityTrigger] IDurableActivityContext context,
            [CosmosDB(
                Connection = "CosmosDBConnection",
                PartitionKey = "/id",
                CreateIfNotExists = true
                )] CosmosClient client,
            ILogger log)
        {
            var database = client.GetDatabase("MeetingMigrationService");
            var container = database.GetContainer("UserMappings");
            try
            {
                await container.DeleteContainerAsync();
                log.LogInformation("Successfully deleted old user mappings container");
            }
            catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                log.LogInformation("[Migration] User mappings container does not exist or already deleted");
            }
        }

        [FunctionName(nameof(UpdateUsersMapping))]
        public static async Task UpdateUsersMapping(
            [ActivityTrigger] UserMapping[] mappingsToUpdate,
            [CosmosDB(
                databaseName: "MeetingMigrationService",
                containerName: "UserMappings",
                Connection = "CosmosDBConnection",
                PartitionKey = "/id",
                CreateIfNotExists = true)] IAsyncCollector<dynamic> mappings,
            ILogger log)
        {
            log.LogInformation("[Migration] Trying to update users mappings");
            foreach (var mappingToUpdate in mappingsToUpdate)
            {
                log.LogInformation("[Migration] Updating mapping {MappingToUpdate}", mappingToUpdate);
                await mappings.AddAsync(new {
                    id = Guid.NewGuid().ToString(),
                    mappingToUpdate.SourceUpn,
                    mappingToUpdate.DestinationUpn
                });
            }
            log.LogInformation("[Migration] Updated users mappings");
        }

        [FunctionName(nameof(ReadMappingFromCsv))]
        public static Task<Dictionary<string, string>> ReadMappingFromCsv(
            [ActivityTrigger] string csv
            )
        {
            var lines = csv.Split(Environment.NewLine);
            // Assert missing key or value
            var incorrectLinesSb = new StringBuilder();
            foreach (var line in lines)
            {
                var upns = line.Split(",");
                if (upns.Length != 2 || string.IsNullOrEmpty(upns[0]) || string.IsNullOrEmpty(upns[1]))
                {
                    var index = Array.IndexOf(lines, line);
                    incorrectLinesSb.AppendLine($"[{index}: ({line})]");
                }
            }
            
            if (incorrectLinesSb.Length > 0)
            {
                throw new InvalidDataException($"Invalid user mapping CSV: \n {incorrectLinesSb}");
            }

            return Task.FromResult(new Dictionary<string, string>(
                csv.Split(Environment.NewLine)
                    .Select(line =>
                    {
                        var upns = line.Split(",");
                        return KeyValuePair.Create(upns[0], upns[1]);
                    })
            ));
        }
    }
}
