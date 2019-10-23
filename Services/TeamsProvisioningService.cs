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
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;

namespace OrchestratedProvisioning.Services
{
    class TeamsProvisioningService
    {
        public async Task<QueueMessage> CreateTeam(QueueMessage message)
        {
            var reader = new TemplateReader();
            var templateString = await reader.Read(message);

            var clientId = ConfigurationManager.AppSettings[AppConstants.KEY_ClientId];
            var builder = PublicClientApplicationBuilder.Create(clientId).WithTenantId(ConfigurationManager.AppSettings[AppConstants.KEY_TenantId]);
            var app = builder.Build();

            var userName = ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningUser];
            var scopes = new string[] { "Group.ReadWrite.All" };

            using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
            {
                var tokenService = new TokenService(app);
                var token = await tokenService.AcquireATokenFromCacheOrUsernamePasswordAsync(scopes, userName, password);

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", token.AccessToken);

                    var body = new StringContent(templateString);
                    body.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                    var response = await client.PostAsync(
                        "https://graph.microsoft.com/beta/teams", body);

                    if (!response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        var responseJson = JObject.Parse(responseContent);
                        var errorMessage = responseJson["error"]["message"].ToString();

                        throw new Exception(errorMessage);
                    }

                    response.EnsureSuccessStatusCode();

                    var operationUrl = response.Headers.Location;
                    var done = false;
                    HttpResponseMessage opResponse = null;
                    while (!done)
                    {
                        await Task.Delay(5000);

                        opResponse = await client.GetAsync(operationUrl);
                        opResponse.EnsureSuccessStatusCode();

                        done = opResponse.Headers.GetValues("status")?.FirstOrDefault<string>() == "succeeded";
                    }

                    string teamId = opResponse?.Headers.GetValues("id")?.FirstOrDefault<string>();
                    //string resultContent = await response.Content.ReadAsStringAsync();

                    message.description = teamId;
                }

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
