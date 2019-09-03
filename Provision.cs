using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using OrchestratedProvisioning.Model;
using OrchestratedProvisioning.Services;

namespace OrchestratedProvisioning
{
    public static class Provision
    {
        [FunctionName("Provision")]
        public static void Run(
            // Input binding is the request queue
            [QueueTrigger(AppConstants.RequestQueueName, Connection = AppConstants.KEY_Storage)]string requestItem,

            // Output binding is the completion queue
            [Queue(AppConstants.CompletionQueueName, Connection = AppConstants.KEY_Storage)] out string completionItem,
            TraceWriter log)
        {
            var completionMessage = new QueueMessage();

            try
            {
                var requestMessage = QueueMessage.NewFromJson(requestItem);
                switch (requestMessage.command)
                {
                    case QueueMessage.Command.provisionModernTeamSite:
                        {
                            var pnpProvisioningService = new PnPTemplateService();
                            completionMessage = pnpProvisioningService.ApplyProvisioningTemplate(requestMessage);
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

            completionItem = completionMessage.Serialize();
            log.Info($"Provision function processed: {completionItem}");
        }
    }
}
