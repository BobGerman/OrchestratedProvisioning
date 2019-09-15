using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Sites;
using OrchestratedProvisioning.Model;
using System;
using System.Configuration;
using System.Threading.Tasks;

namespace OrchestratedProvisioning.Services
{
    class PnPTemplateService
    {
        public async Task<QueueMessage> ProvisionWithTemplate(QueueMessage message)
        {
            var rootSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_RootSiteUrl];
            string newSiteUrl = null;

            message.resultCode = QueueMessage.ResultCode.unknown;

            try
            {
                // Part 1: Create the new site
                newSiteUrl = await CreateSiteAsync(message, rootSiteUrl, newSiteUrl);

                // Part 2: Get the provisioning template
                ProvisioningTemplate provisioningTemplate = null;
                provisioningTemplate = await GetProvisioningTemplateAsync(message);

                // Part 3: Apply the provisioning template

                await ApplyProvisioningTemplateAsync(message, newSiteUrl, provisioningTemplate);

                message.resultCode = QueueMessage.ResultCode.success;
            }
            catch (Exception ex)
            {
                message.resultCode = QueueMessage.ResultCode.failure;
                message.resultMessage = ex.Message;
            }

            return message;
        }

        private static async Task<string> CreateSiteAsync(QueueMessage message, string rootSiteUrl, string newSiteUrl)
        {
            await CsomProviderService.GetContextAsync(rootSiteUrl, (async (ctx) =>
            {
                var siteContext = await ctx.CreateSiteAsync(
                    new TeamSiteCollectionCreationInformation
                    {
                        Alias = message.alias, // Mandatory
                        DisplayName = !String.IsNullOrEmpty(message.displayName) ? message.displayName : message.alias, // Mandatory
                        Description = message.description, // Optional
                                                           //                            Classification = "classification", // Optional
                        IsPublic = true, // Optional, default true
                    });

                var web = siteContext.Web;
                siteContext.Load(web, w => w.Title, w => w.ServerRelativeUrl);
                await siteContext.ExecuteQueryRetryAsync();

                // Combine the root and relative URL of the new site
                newSiteUrl = (new Uri((new Uri(rootSiteUrl)), web.ServerRelativeUrl)).AbsoluteUri;

                message.resultMessage = $"Created {web.Title} at {newSiteUrl}";
                message.displayName = web.Title;
                message.requestId = web.ServerRelativeUrl;
            }));
            return newSiteUrl;
        }

        private static async Task<ProvisioningTemplate> GetProvisioningTemplateAsync(QueueMessage message)
        {
            var templateSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_TemplateSiteUrl];
            ProvisioningTemplate provisioningTemplate = null;

            await CsomProviderService.GetContextAsync(templateSiteUrl, (async (ctx) =>
            {
                // Thanks Anuja Bhojani for posting a sample of how to use this
                // http://anujabhojani.blogspot.com/2017/11/pnp-example-of-xmlsharepointtemplatepro.html

                XMLSharePointTemplateProvider provider = new XMLSharePointTemplateProvider(ctx, templateSiteUrl, ConfigurationManager.AppSettings[AppConstants.KEY_TemplateLibrary]);

                provisioningTemplate = provider.GetTemplate(message.template);

                message.resultMessage = $"Retrieved template {message.template} from {templateSiteUrl}";

            }));
            return provisioningTemplate;
        }

        private static async Task ApplyProvisioningTemplateAsync(QueueMessage message, string newSiteUrl, ProvisioningTemplate provisioningTemplate)
        {
            await CsomProviderService.GetContextAsync(newSiteUrl, (async (ctx) =>
            {
                Web web = ctx.Web;
                ctx.Load(web);
                await ctx.ExecuteQueryRetryAsync();

                // Is there an async version?
                web.ApplyProvisioningTemplate(provisioningTemplate);

                message.resultMessage = $"Applied provisioning template {message.template} to {newSiteUrl}";

            }));
        }
    }
}
