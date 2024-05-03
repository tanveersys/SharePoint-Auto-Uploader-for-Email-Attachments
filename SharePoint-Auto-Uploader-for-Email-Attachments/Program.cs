using Azure.Identity;
using Microsoft.Graph;

namespace SharePoint_Auto_Uploader_for_Email_Attachments
{
    // https://damienbod.com/2023/01/16/implementing-secure-microsoft-graph-application-clients-in-asp-net-core/
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string clientId = "app_client_id";
            string tenantId = "azure_ad_tenant_id";
            string clientSecret = "app_client_secret";
            string sharePointSiteUrl = "https://sharepoint_site";
            string targetFolder = "/Shared Documents/MyAttachments";
            string userEmail = "tanveer.sys@gmail.com";
            string subjectFilter = "Contact xy1223";

            try
            {
                var graphClient = GetGraphClientWithClientSecretCredential(clientId, tenantId, clientSecret);
                var processor = new EmailAttachmentProcessor(graphClient);

                await processor.CheckEmails(userEmail, subjectFilter, sharePointSiteUrl, targetFolder);
                Console.WriteLine("Finished processing emails.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        private static GraphServiceClient GetGraphClientWithClientSecretCredential(string clientId, string tenantId, string clientSecret)
        {
            string[] scopes = new[] { "https://graph.microsoft.com/.default" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

            return new GraphServiceClient(clientSecretCredential, scopes);
        }
    }
}
