using System.Configuration;
using System.Security;
using System.Text;
using System.Threading;
using Microsoft.SharePoint.Client;
using OrchestratedProvisioning.Model;

namespace OrchestratedProvisioning.Services
{
    class TemplateReader
    {
        public string Read(string templateName)
        {
            var templateSiteUrl = ConfigurationManager.AppSettings[AppConstants.KEY_TemplateSiteUrl];
            var templateLibrary = ConfigurationManager.AppSettings[AppConstants.KEY_TemplateLibrary];
            var userName = ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningUser];

            using (var ctx = new ClientContext(templateSiteUrl))
            {
                using (var password = GetSecureString(ConfigurationManager.AppSettings[AppConstants.KEY_ProvisioningPassword]))
                {
                    ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                    ctx.RequestTimeout = Timeout.Infinite;
                    var resultBuilder = new StringBuilder();

                    List list = ctx.Web.Lists.GetByTitle(templateLibrary);
                    ctx.Load(list);
                    ctx.ExecuteQueryRetry();

                    var folder = list.RootFolder;
                    var files = folder.Files;
                    ctx.Load(files);
                    ctx.ExecuteQueryRetry();

                    foreach (var file in files)
                    {
                        if (file.Name.ToLower() == templateName.ToLower())
                        {
                            FileInformation fileInformation = File.OpenBinaryDirect(ctx, (string)file.ServerRelativeUrl);
                            using (System.IO.StreamReader sr = new System.IO.StreamReader(fileInformation.Stream))
                            {
                                // Read the stream to a string, and write the string to the console.
                                string line = sr.ReadToEnd();
                                resultBuilder.AppendLine(line);
                            }
                        }
                    }

                    return resultBuilder.ToString();
                }
            }
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
