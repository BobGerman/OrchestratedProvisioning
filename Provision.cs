using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using OrchestratedProvisioning.Model;
using OrchestratedProvisioning.Services;

namespace OrchestratedProvisioning
{
    public static class Provision
    {
        [FunctionName("Provision")]
        public static async Task Run(
            // Input binding is the request queue
            [QueueTrigger(Constants.RequestQueueName, Connection = SettingKey.Storage)]string requestItem,

            // Output binding is the completion queue
            [Queue(Constants.CompletionQueueName, Connection = SettingKey.Storage)] IAsyncCollector<string> completionItem,
            TraceWriter log)
        {
            var completionMessage = new QueueMessage();

            try
            {
                var requestMessage = QueueMessage.NewFromJson(requestItem);
                if (string.IsNullOrEmpty(requestMessage.template))
                {
                    throw new Exception("Empty template name");
                }
                
                switch (requestMessage.command)
                {
                    case QueueMessage.Command.provisionModernTeamSite:
                        {
                            var pnpProvisioningService = new PnPTemplateService();
                            completionMessage = await pnpProvisioningService.ProvisionWithTemplateAsync(requestMessage);
                            break;
                        }

                    case QueueMessage.Command.applyProvisioningTemplate:
                        {
                            var pnpProvisioningService = new PnPTemplateService();
                            completionMessage = await pnpProvisioningService.ApplyProvisioningTemplateAsync(requestMessage);
                            break;
                        }

                    case QueueMessage.Command.createTeam:
                        {
                            var teamsProvisioningService = new TeamsProvisioningService();
                            completionMessage = await teamsProvisioningService.CreateTeam(requestMessage);
                            break;
                        }

                    default:
                        {
                            completionMessage.resultCode = QueueMessage.ResultCode.failure;
                            completionMessage.resultMessage = "Unknown command";
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                completionMessage.resultCode = QueueMessage.ResultCode.failure;
                completionMessage.resultMessage = ex.Message;
            }

            await completionItem.AddAsync(completionMessage.Serialize());
            log.Info($"Provision function processed: {completionItem}");
        }
    }
}
