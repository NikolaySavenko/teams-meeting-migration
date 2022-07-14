using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Services;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Primitives;
using Newtonsoft.Json;
using TeamsMigrationFunction.MailboxOrchestration;

namespace TeamsMigrationFunction.EmailSending
{
    public static class EmailSender
    {
        public static DurableHttpRequest BuildEmailRequest(EmailMessage email)
        {
            var emailEndpoint = Environment.GetEnvironmentVariable("EmailEndpoint");
            var headers = new Dictionary<string, StringValues>()
            {
                {"Content-Type", "application/json"},
                {"Location", ""}
            };
             
            // var content = "{ \"to\": \"ns@o365hq.com\",\"from\": \"MMS\",\"subject\": \"test\", \"body\": \"some another\"}";
            var content = JsonConvert.SerializeObject(email);
            return new DurableHttpRequest(HttpMethod.Post, new Uri(emailEndpoint), headers, content, asynchronousPatternEnabled: false);
        }
    }
    public record EmailMessage(string from, string to, string subject, string body);
}
