namespace OrchestratedProvisioning.Model
{
    // Various constants used in the application
    static public class Constants
    { 
        // Queue names
        public const string RequestQueueName = "provisioning-request-queue";
        public const string CompletionQueueName = "provisioning-completion-queue";
    }

    // Names of settings in the Azure app service or settings.json
    static public class SettingKey
    {
        public const string Storage = "AzureWebJobsStorage";
        public const string ProvisioningUser = "ProvisioningServiceUser";
        public const string ProvisioningPassword = "ProvisioningServicePassword";
        public const string TenantId = "TenantId";
        public const string ClientId = "ClientId";
        public const string RootSiteUrl = "RootSiteUrl";
        public const string TemplateSiteUrl = "TemplateSiteUrl";
        public const string TemplateLibrary = "TemplateLibrary";
    }
}
