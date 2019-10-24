using Microsoft.SharePoint.Client;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace OrchestratedProvisioning.Services
{
    public static class PnPContextProvider
    {
        public delegate Task Callback(ClientContext ctx);
        public static async Task WithContextAsync(string siteUrl, Callback callback)
        {
            var userName = Settings.GetString(Settings.Key.ProvisioningUser);

            using (var ctx = new ClientContext(siteUrl))
            {
                using (var password = GetSecureString(Settings.Key.ProvisioningPassword))
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
