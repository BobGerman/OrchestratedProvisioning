using Microsoft.SharePoint.Client;
using OrchestratedProvisioning.Model;
using System.Configuration;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace OrchestratedProvisioning.Services
{
    public static class CsomProviderService
    {
        public delegate Task Callback(ClientContext ctx);
        public static async Task GetContextAsync(string siteUrl, Callback callback)
        {
            var userName = ConfigurationManager.AppSettings[SettingKey.ProvisioningUser];

            using (var ctx = new ClientContext(siteUrl))
            {
                using (var password = GetSecureString(ConfigurationManager.AppSettings[SettingKey.ProvisioningPassword]))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                    ctx.RequestTimeout = Timeout.Infinite;

                    await callback(ctx);
                }
            }
        }

        private static SecureString GetSecureString(string plaintext)
        {
            var result = new SecureString();
            foreach (var c in plaintext)
            {
                result.AppendChar(c);
            }
            return result;
        }
    }
}
