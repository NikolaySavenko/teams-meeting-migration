using System;
using System.Threading.Tasks;

namespace TeamsMigrationFunction.UserConfiguration
{
    public interface IUserConfiguration
    {
        public Task SetMailboxStartTime(string start);
        public Task<string> GetMailboxStartTime();
    }
}
