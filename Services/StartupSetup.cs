using System.Globalization;
using System.Net.Http.Headers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;


namespace Services
{
    public static class StartupSetup
    {
        private const string ApiUrl = "https://graph.microsoft.com/";

        private const string Instance = "https://login.microsoftonline.com/{0}";

        public static void AddTenantGraphClient(this IServiceCollection services, string tenant, string clientId, string clientSecret) => 
        services.AddSingleton(s =>
        {
            var authority = string.Format(CultureInfo.InvariantCulture, Instance, tenant);
            var app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri(authority))
                .Build()
                .AddInMemoryTokenCache();
            var scopes = new string[] { $"{ApiUrl}.default" };
            var graphClient = new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                var result = await app.AcquireTokenForClient(scopes)
                    .ExecuteAsync();
        
                // Add the access token in the Authorization header of the API request.
                requestMessage.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", result.AccessToken);
            }));
            var logger = s.GetService<ILogger<TenantGraphClient>>();
            
            return new TenantGraphClient(graphClient);
        });
    }
}
