using System.Configuration;
using System.Security;

namespace OrchestratedProvisioning.Services
{
    class Settings
    {
        // Names of settings in the Azure app service or settings.json
        static public class Key
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

        public static string GetString(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        public static SecureString GetSecureString(string key)
        {
            var plaintext = ConfigurationManager.AppSettings[key];

            var result = new SecureString();
            foreach (var c in plaintext)
            {
                result.AppendChar(c);
            }
            return result;
        }
    }
}
