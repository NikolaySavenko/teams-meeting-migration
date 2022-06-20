using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace TeamsMigrationFunction.UserMapping
{
    public interface IUserMapper
    {
        Task AddUserMapping(KeyValuePair<string, string> mapping);
        Task<string> GetUserDestinationUpn(string sourceUpn);
        Task<Attendee> GetDestinationAttendee(Attendee attendee);
        Task RecreateUsersMappings(IDictionary<string, string> mapping);
    }
}
