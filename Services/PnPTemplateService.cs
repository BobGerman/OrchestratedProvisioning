using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OrchestratedProvisioning.Model;

namespace OrchestratedProvisioning.Services
{
    class PnPTemplateService
    {
        public async Task<QueueMessage> ApplyProvisioningTemplate (QueueMessage request)
        {
            var userName = ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningUser];
            var rootSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_RootSiteUrl];

            var result = request;

            using (var ctx = new ClientContext(rootSiteUrl))
            {
                using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                    ctx.RequestTimeout = Timeout.Infinite;

                    var web = ctx.Web;
                    ctx.Load(web, w => w.Title);
                    await ctx.ExecuteQueryRetryAsync();

                    result.resultCode = QueueMessage.ResultCode.success;
                    result.resultMessage = web.Title;
                    result.displayName = web.Title;
                }
            }

            return result;
        }

        // Converts string to secure string for use in CSOM
        // Not to be used in untrusted hosting env't as string is still handled
        // in the clear
        private SecureString GetSecureString(string plaintext)
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
