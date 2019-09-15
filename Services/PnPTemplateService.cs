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
using OfficeDevPnP.Core.Sites;
using OrchestratedProvisioning.Model;

namespace OrchestratedProvisioning.Services
{
    class PnPTemplateService
    {
        public async Task<QueueMessage> ProvisionWithTemplate(QueueMessage request)
        {
            var userName = ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningUser];
            var rootSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_RootSiteUrl];
            var templateSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_TemplateSiteUrl];
            string newSiteUrl = null;

            var result = request;
            result.resultCode = QueueMessage.ResultCode.unknown;

            try
            {
                // Part 1: Create the new site
                using (var ctx = new ClientContext(rootSiteUrl))
                {
                    using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
                    {
                        ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                        ctx.RequestTimeout = Timeout.Infinite;

                        var siteContext = await ctx.CreateSiteAsync(
                            new TeamSiteCollectionCreationInformation
                            {
                                Alias = request.alias, // Mandatory
                            DisplayName = request.displayName, // Mandatory
                            Description = request.description, // Optional
                                                               //                            Classification = "classification", // Optional
                            IsPublic = true, // Optional, default true
                        });

                        var web = siteContext.Web;
                        siteContext.Load(web, w => w.Title, w => w.ServerRelativeUrl);
                        await siteContext.ExecuteQueryRetryAsync();

                        // Combine the root and relative URL of the new site
                        newSiteUrl = (new Uri((new Uri(rootSiteUrl)), web.ServerRelativeUrl)).AbsoluteUri;

                        result.resultMessage = $"Created {web.Title} at {newSiteUrl}";
                        result.displayName = web.Title;
                        result.requestId = web.ServerRelativeUrl;

                    }
                }

                // Part 2: Get the provisioning template
                ProvisioningTemplate provisioningTemplate = null;
                using (var ctx = new ClientContext(templateSiteUrl))
                {
                    using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
                    {
                        ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                        ctx.RequestTimeout = Timeout.Infinite;

                        // Thanks Anuja Bhojani for posting a sample of how to use this
                        // http://anujabhojani.blogspot.com/2017/11/pnp-example-of-xmlsharepointtemplatepro.html

                        XMLSharePointTemplateProvider provider = new XMLSharePointTemplateProvider(ctx, templateSiteUrl, ConfigurationManager.AppSettings[AppConstants.KEY_TemplateLibrary]);

                        // Is there an async version?
                        provisioningTemplate = provider.GetTemplate(request.template);

                        result.resultMessage = $"Retrieved template {request.template} from {templateSiteUrl}";
                    }
                }

                // Part 3: Apply the provisioning template
                using (var ctx = new ClientContext(newSiteUrl))
                {
                    using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
                    {
                        ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                        ctx.RequestTimeout = Timeout.Infinite;

                        Web web = ctx.Web;
                        ctx.Load(web);
                        await ctx.ExecuteQueryRetryAsync();

                        // Is there an async version?
                        web.ApplyProvisioningTemplate(provisioningTemplate);

                        result.resultMessage = $"Applied provisioning template {request.template} to {newSiteUrl}";
                    }
                }

                result.resultCode = QueueMessage.ResultCode.success;
            }
            catch (Exception ex)
            {
                result.resultCode = QueueMessage.ResultCode.failure;
                result.resultMessage = ex.Message;
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
