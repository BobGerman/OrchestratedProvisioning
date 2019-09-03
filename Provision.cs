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
                var reader = new TemplateReader();
                if (string.IsNullOrEmpty(requestMessage.template))
                {
                    throw new Exception("Empty template name");
                }
                var templateString = reader.Read(requestMessage.template);

                switch (requestMessage.command)
                {
                    case QueueMessage.Command.provisionModernTeamSite:
                        {
                            var pnpProvisioningService = new PnPTemplateService();
                            completionMessage = pnpProvisioningService.ApplyProvisioningTemplate(requestMessage);
                            completionMessage.resultMessage = templateString;
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
