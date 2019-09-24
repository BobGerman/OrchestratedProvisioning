using OrchestratedProvisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Graph;
using Microsoft.Identity.Client;
using System.Configuration;
using System.Security;

namespace OrchestratedProvisioning.Services
{
    class TeamsProvisioningService
    {
        public async Task<QueueMessage> CreateTeam(QueueMessage message)
        {
            var clientId = ConfigurationManager.AppSettings[AppConstants.KEY_ClientId];
            var builder = PublicClientApplicationBuilder.Create(clientId);
            var app = builder.Build();

            var userName = ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningUser];
            var scopes = new string[] { "User.Read", "User.ReadBasic.All" };

            using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
            {
                var tokenService = new TokenService(app);
                var token = await tokenService.AcquireATokenFromCacheOrUsernamePasswordAsync(scopes, userName, password);

                message.description = token.AccessToken;

            }




            return message;
        }

        private static SecureString GetSecureString(string plaintext)
        {
            var result = new SecureString();
            foreach (var c in plaintext)
            {
                result.AppendChar(c);
            }
            return result;
        }

    }
}
