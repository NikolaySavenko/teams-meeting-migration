using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Graph;
using Newtonsoft.Json;
using KeyValuePair = System.Collections.Generic.KeyValuePair;

namespace TeamsMigrationFunction.UserMapping
{
    public class UserMapper : IUserMapper
    {
        [JsonProperty("contacts")]
        private IDictionary<string, string> UsersMappings { get; set; }

        public Task AddUserMapping(KeyValuePair<string, string> mapping)
        {
            UsersMappings.Add(mapping);
            return Task.CompletedTask;
        }

        public Task<string> GetUserDestinationUpn(string sourceUpn)
        {
            var success = UsersMappings.TryGetValue(sourceUpn, out var destinationUpn);
            return Task.FromResult(success ? destinationUpn : sourceUpn);
        }

        public async Task<Attendee> GetDestinationAttendee(Attendee attendee)
        {
            attendee.EmailAddress.Address = await GetUserDestinationUpn(attendee.EmailAddress.Address);
            // attendee.Status = new ResponseStatus
            // {
            //     Response = ResponseType.NotResponded
            // };
            return attendee;
        }

        public Task RecreateUsersMappings(IDictionary<string, string> mapping)
        {
            UsersMappings = mapping;
            return Task.CompletedTask;
        }
        
        [FunctionName(nameof(UserMapper))]
        public static Task Run([EntityTrigger] IDurableEntityContext ctx)
            => ctx.DispatchAsync<UserMapper>();
        
    }
}
