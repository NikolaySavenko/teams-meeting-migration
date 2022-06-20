using System;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Services;
using TeamsMigrationFunction;

[assembly: FunctionsStartup(typeof(Startup))]

namespace TeamsMigrationFunction
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            builder.Services.AddHttpClient();
            // TODO change target tenant creds
            builder.Services.AddTenantGraphClient(
                Environment.GetEnvironmentVariable("SourceTenantId") ?? throw new ArgumentNullException("SourceTenantId","No SourceTenantId provided"),
                Environment.GetEnvironmentVariable("SourceClientId") ?? throw new ArgumentNullException("SourceClientId","No SourceClientId provided"),
                Environment.GetEnvironmentVariable("SourceClientSecret") ?? throw new ArgumentNullException("SourceClientSecret", "No SourceClientSecret provided")
            );
        }
    }
}
