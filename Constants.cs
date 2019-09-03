namespace OrchestratedProvisioning
{
    static public class Constants
    {
        // Keys to settings in Azure or local.settings.json
        public const string KEY_Storage = "AzureWebJobsStorage";
        public const string KEY_ProvisioningUser = "ProvisioningServiceUser";
        public const string KEY_ProvisioningPassword = "ProvisioningServicePassword";
        public const string KEY_RootSiteUrl = "RootSiteUrl";

        // Queues
        public const string RequestQueueName = "provisioning-request-queue";
        public const string CompletionQueueName = "provisioning-completion-queue";
    }
}
