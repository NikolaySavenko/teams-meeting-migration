using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Newtonsoft.Json;

namespace TeamsMigrationFunction.UserConfiguration
{
    public class UserConfiguration : IUserConfiguration
    {
        public const string UserAdditionalFieldName = "mailboxStartTime";
        
        [JsonProperty("mailboxStartTime")]
        private string _mailboxStartTime;
        
        public Task SetMailboxStartTime(string start)
        {
            _mailboxStartTime = start;
            return Task.CompletedTask;
        }

        public Task<string> GetMailboxStartTime()
        {
            return Task.FromResult(_mailboxStartTime);
        }

        [FunctionName(nameof(UserConfiguration))]
        public static Task Run([EntityTrigger] IDurableEntityContext ctx)
            => ctx.DispatchAsync<UserConfiguration>();
    }
}
