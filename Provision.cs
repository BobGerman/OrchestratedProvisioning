using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using OrchestratedProvisioning.Services;

namespace OrchestratedProvisioning
{
    public static class Provision
    {
        [FunctionName("Provision")]
        public static void Run(
            // Input binding is the request queue
            [QueueTrigger(Constants.RequestQueueName, Connection = Constants.KEY_Storage)]string requestItem,

            // Output binding is the completion queue
            [Queue(Constants.CompletionQueueName, Connection = Constants.KEY_Storage)] out string completionItem,
            TraceWriter log)
        {

            var pnpProvisioningService = new PnPTemplateService();
            completionItem = pnpProvisioningService.ApplyProvisioningTemplate(requestItem);

            log.Info($"Provision function processed: {requestItem}");
        }
    }
}
