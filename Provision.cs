using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;

namespace OrchestratedProvisioning
{
    public static class Provision
    {
        [FunctionName("Provision")]
        public static void Run(
            // Input binding is the request queue
            [QueueTrigger(Constants.RequestQueueName, Connection = Constants.SettingsKey4Storage)]string requestItem,

            // Output binding is the completion queue
            [Queue(Constants.CompletionQueueName, Connection = Constants.SettingsKey4Storage)] out string completionItem,
            TraceWriter log)
        {
            completionItem = requestItem;
            log.Info($"Provision function processed: {requestItem}");
        }
    }
}
