using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace Extensions.GraphAPI
{

    public static class GraphAPIServiceExtension
    {
        public static void AddGraphAPIService(this IServiceCollection services, IConfiguration configuration)
        {

            SecureString secure = new SecureString();
            configuration["GraphAPI:SystemUser:Password"].ToCharArray().ToList().ForEach(c => secure.AppendChar(c));
            secure.MakeReadOnly();
            services.AddSingleton<GraphAuthenticator>((container) =>
            {
                var logger = container.GetRequiredService<ILogger<GraphAuthenticator>>();

                return new GraphAuthenticator(
                configuration["AzureAd:ClientId"],
                configuration["AzureAd:ClientSecret"],
                configuration["AzureAd:TenantId"],
                configuration["AzureAd:RedirectURL"],
                configuration["GraphAPI:Scope"],
                configuration["GraphAPI:SystemUser:Name"],
                secure)
                {
                    Logger = logger
                };
            });

            services.AddTransient<GraphService>();
        }
    }
}
