using System.Configuration;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OrchestratedProvisioning.Model;

namespace OrchestratedProvisioning.Services
{
    class TemplateReader
    {
        public async Task<string> Read(QueueMessage message)
        {
            var templateSiteUrl = ConfigurationManager.AppSettings[SettingKey.TemplateSiteUrl];
            var templateLibrary = ConfigurationManager.AppSettings[SettingKey.TemplateLibrary];
            var userName = ConfigurationManager.AppSettings[SettingKey.ProvisioningUser];

            using (var ctx = new ClientContext(templateSiteUrl))
            {
                using (var password = GetSecureString(ConfigurationManager.AppSettings[SettingKey.ProvisioningPassword]))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                    ctx.RequestTimeout = Timeout.Infinite;
                    var resultBuilder = new StringBuilder();

                    List list = ctx.Web.Lists.GetByTitle(templateLibrary);
                    ctx.Load(list);
                    await ctx.ExecuteQueryRetryAsync();

                    var folder = list.RootFolder;
                    var files = folder.Files;
                    ctx.Load(files);
                    await ctx.ExecuteQueryRetryAsync();

                    foreach (var file in files)
                    {
                        if (file.Name.ToLower() == message.template.ToLower())
                        {
                            FileInformation fileInformation = File.OpenBinaryDirect(ctx, (string)file.ServerRelativeUrl);
                            using (System.IO.StreamReader sr = new System.IO.StreamReader(fileInformation.Stream))
                            {
                                string line = await sr.ReadToEndAsync();
                                resultBuilder.AppendLine(line);
                            }
                        }
                    }

                    var resultJson = JObject.Parse(resultBuilder.ToString());
                    AddOrReplaceJsonProperty(resultJson, "visibility", message.isPublic ? "public" : "private");
                    AddOrReplaceJsonProperty(resultJson, "displayName", message.displayName);
                    AddOrReplaceJsonProperty(resultJson, "description", message.description);

                    var ownerId = await GetUserIdFromUpnAsync(message.owner);
                    var ownersJson = new JArray();
                    ownersJson.Add($"https://graph.microsoft.com/beta/users('{ownerId}')");

                    AddOrReplaceJsonProperty(resultJson, "owners@odata.bind", ownersJson);


                    var result = resultJson.ToString();

                    return result;
                }
            }
        }

        private async Task<string> GetUserIdFromUpnAsync (string upn)
        {
            return "cda79e6a-d052-4eb8-b192-7e039ad0ed73";
        }

        private void AddOrReplaceJsonProperty(JObject j, string key, string value)
        {
            AddOrReplaceJsonProperty(j, key, JToken.FromObject(value));
        }

        private void AddOrReplaceJsonProperty(JObject j, string key, JToken value)
        {
            if (j.ContainsKey(key))
            {
                j.Remove(key);
            }
            j.Add(key, value);
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
