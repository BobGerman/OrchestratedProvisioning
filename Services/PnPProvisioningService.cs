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
    class PnPProvisioningService
    {
        #region Service methods
        public async Task<QueueMessage> ProvisionWithTemplateAsync(QueueMessage message)
        {
            string newSiteUrl = null;

            message.resultCode = QueueMessage.ResultCode.unknown;

            try
            {
                // Part 1: Create the new site
                newSiteUrl = await CreateSiteAsync(message,  newSiteUrl);

                // Part 2: Get the provisioning template
                ProvisioningTemplate provisioningTemplate = null;
                provisioningTemplate = await GetProvisioningTemplateAsync(message);

                // Part 3: Apply the provisioning template

                await ApplyProvisioningTemplateAsync(message, newSiteUrl, provisioningTemplate);

                message.resultCode = QueueMessage.ResultCode.succeeded;
            }
            catch (Exception ex)
            {
                message.resultCode = QueueMessage.ResultCode.failure;
                message.resultMessage = ex.Message;
            }

            return message;
        }

        public async Task<QueueMessage> ApplyProvisioningTemplateAsync(QueueMessage message)
        {
            message.resultCode = QueueMessage.ResultCode.unknown;
            var rootSiteUrl = Settings.GetString(Settings.Key.RootSiteUrl);


            try
            {
                // Part 1: Update message from site
                var siteUrl = await GetSiteInfoAsync(message);

                // Part 2: Get the provisioning template
                ProvisioningTemplate provisioningTemplate = null;
                provisioningTemplate = await GetProvisioningTemplateAsync(message);

                // Part 3: Apply the provisioning template

                await ApplyProvisioningTemplateAsync(message, siteUrl, provisioningTemplate);

                message.resultCode = QueueMessage.ResultCode.succeeded;
            }
            catch (Exception ex)
            {
                message.resultCode = QueueMessage.ResultCode.failure;
                message.resultMessage = ex.Message;
            }

            return message;
        }

        #endregion

        #region Private methods

        private static async Task<string> CreateSiteAsync(QueueMessage message, string newSiteUrl)
        {
            var rootSiteUrl = Settings.GetString(Settings.Key.RootSiteUrl);
            await PnPContextProvider.WithContextAsync(rootSiteUrl, (async (ctx) =>
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

        private static async Task<string> GetSiteInfoAsync(QueueMessage message)
        {
            var rootSiteUrl = Settings.GetString(Settings.Key.RootSiteUrl);
            var siteUrl = (new Uri((new Uri(rootSiteUrl)), $"/sites/{message.alias}")).AbsoluteUri;

            await PnPContextProvider.WithContextAsync(siteUrl, (async (ctx) =>
            {
                var web = ctx.Web;
                ctx.Load(web, w => w.Title, w => w.ServerRelativeUrl);
                await ctx.ExecuteQueryRetryAsync();

                message.resultMessage = $"Found {web.Title} at {siteUrl}";
                message.displayName = web.Title;
                message.requestId = web.ServerRelativeUrl;
            }));
            return siteUrl;
        }

        private static async Task<ProvisioningTemplate> GetProvisioningTemplateAsync(QueueMessage message)
        {
            var templateSiteUrl = Settings.GetString(Settings.Key.TemplateSiteUrl);
            ProvisioningTemplate provisioningTemplate = null;

            await PnPContextProvider.WithContextAsync(templateSiteUrl, (async (ctx) =>
            {
                // Thanks Anuja Bhojani for posting a sample of how to use this
                // http://anujabhojani.blogspot.com/2017/11/pnp-example-of-xmlsharepointtemplatepro.html

                XMLSharePointTemplateProvider provider = new XMLSharePointTemplateProvider(ctx, templateSiteUrl, Settings.GetString(Settings.Key.TemplateLibrary));

                provisioningTemplate = provider.GetTemplate(message.template);

                message.resultMessage = $"Retrieved template {message.template} from {templateSiteUrl}";

            }));
            return provisioningTemplate;
        }

        private static async Task ApplyProvisioningTemplateAsync(QueueMessage message, string newSiteUrl, ProvisioningTemplate provisioningTemplate)
        {
            await PnPContextProvider.WithContextAsync(newSiteUrl, (async (ctx) =>
            {
                Web web = ctx.Web;
                ctx.Load(web);
                await ctx.ExecuteQueryRetryAsync();

                // Is there an async version?
                web.ApplyProvisioningTemplate(provisioningTemplate);

                message.resultMessage = $"Applied provisioning template {message.template} to {newSiteUrl}";

            }));
        }

        #endregion
    }
}
