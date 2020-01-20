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
                var resultCode = QueueMessage.ResultCode.unknown;
                var ownerId = responseJson["id"]?.ToString();

                var templateString = await reader.Read(message, ownerId);

                (responseJson, resultCode) = await graphClient.PostTeamsAsyncOperation(authResult.AccessToken, "https://graph.microsoft.com/beta/teams", templateString);

                message.groupId = responseJson["id"]?.ToString();
                message.resultMessage = responseJson["status"]?.ToString();
                message.resultValue = responseJson["Value"]?.ToString();
                message.resultCode = resultCode;
            });
            return message;
        }


    }
}
