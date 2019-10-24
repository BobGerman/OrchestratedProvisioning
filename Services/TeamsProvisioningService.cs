using OrchestratedProvisioning.Model;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace OrchestratedProvisioning.Services
{
    class TeamsProvisioningService
    {
        public async Task<QueueMessage> CreateTeam(QueueMessage message)
        {
            var reader = new TeamsTemplateReader();
            var templateString = await reader.Read(message);

            await MSGraphTokenProvider.WithAuthResult(async(AuthenticationResult authResult) =>
            {
                var graphClient = new MSGraphClient();
                string url = "https://graph.microsoft.com/beta/teams";
                var responseJson = await graphClient.PostTeamsAsyncOperation(authResult.AccessToken, url, templateString);

                var teamId = responseJson["id"]?.ToString();
                message.description = teamId;
            });
            return message;
        }


    }
}
