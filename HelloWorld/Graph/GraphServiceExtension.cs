using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace HelloWorld.Graph
{

    public static class GraphAPIServiceExtension
    {
        public static void AddGraphAPIService(this IServiceCollection services, IConfiguration configuration)
        {

            SecureString secure = new SecureString();
            configuration["GraphAPI:SystemUser:Password"].ToCharArray().ToList().ForEach(c => secure.AppendChar(c));
            secure.MakeReadOnly();

            services.AddSingleton<GraphAuthenticator>(new GraphAuthenticator(configuration["AzureAd:ClientId"],
                configuration["AzureAd:ClientSecret"],
                configuration["AzureAd:TenantId"],
                configuration["AzureAd:RedirectURL"],
                configuration["GraphAPI:Scope"],
                configuration["GraphAPI:SystemUser:Name"],
                secure));

            services.AddTransient<GraphService>();
        }
    }
}
