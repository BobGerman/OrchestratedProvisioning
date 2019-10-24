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

            await MSGraphTokenProvider.WithAuthResult(async(AuthenticationResult authResult) =>
            {
                var graphClient = new MSGraphClient();

                var responseJson = await graphClient.Get(authResult.AccessToken, $"https://graph.microsoft.com/v1.0/users/{message.owner}");
                var ownerId = responseJson["id"]?.ToString();

                var templateString = await reader.Read(message, ownerId);

                responseJson = await graphClient.PostTeamsAsyncOperation(authResult.AccessToken, "https://graph.microsoft.com/beta/teams", templateString);

                var teamId = responseJson["id"]?.ToString();
                message.description = teamId;
            });
            return message;
        }


    }
}
