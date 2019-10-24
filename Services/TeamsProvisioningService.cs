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
            var reader = new TeamsTemplateReader();
            var templateString = await reader.Read(message);

            await MSGraphTokenProvider.WithAuthResult(async(AuthenticationResult token) =>
            {
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

                    var operationUrl = "https://graph.microsoft.com/beta" + response.Headers.Location;
                    var done = false;
                    HttpResponseMessage opResponse = null;
                    var teamId = string.Empty;
                    while (!done)
                    {
                        await Task.Delay(5000);

                        opResponse = await client.GetAsync(operationUrl);
                        if (!opResponse.IsSuccessStatusCode)
                        {
                            var responseContent = await opResponse.Content.ReadAsStringAsync();
                            var responseJson = JObject.Parse(responseContent);
                            var errorMessage = responseJson["error"]["message"].ToString();

                            throw new Exception(errorMessage);
                        }

                        var opResponseContent = await opResponse.Content.ReadAsStringAsync();
                        var opResponseJson = JObject.Parse(opResponseContent);
                        done = opResponseJson["status"]?.ToString() == "succeeded";
                        teamId = opResponseJson["id"]?.ToString();
                    }

                    message.description = teamId;
                }

            });
            return message;
        }


    }
}
