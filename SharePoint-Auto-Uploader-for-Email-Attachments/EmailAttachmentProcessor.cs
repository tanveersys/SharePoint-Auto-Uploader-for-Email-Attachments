using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Linq;

namespace SharePoint_Auto_Uploader_for_Email_Attachments
{
    public class EmailAttachmentProcessor
    {
        private readonly GraphServiceClient _graphClient;

        public EmailAttachmentProcessor(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task CheckEmails(string userEmail, string subjectFilter, string sharePointSiteUrl, string targetFolder)
        {
            var filterParams = new List<string>();
            filterParams.Add($"user.mail eq '{userEmail}'");

            if (!string.IsNullOrEmpty(subjectFilter))
            {
                filterParams.Add($"subject eq '{subjectFilter}'");
            }

            var filter = string.Join(" and ", filterParams);


            var messages = await _graphClient.Users[userEmail].Messages.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Filter = filter;
                requestConfiguration.QueryParameters.Top = 100;
            });


            if (!messages.any())
            {
                Console.WriteLine("No emails found with the specified subject.");
                return;
            }

            foreach (var message in messages)
            {
                if (message.HasAttachments)
                {
                    Console.WriteLine($"Processing email: {message.Id}");
                    foreach (var attachment in message.Attachments)
                    {
                        await DownloadAndUploadAttachments(message.Id, userEmail, sharePointSiteUrl, targetFolder);
                        Console.WriteLine($"Uploaded attachment: {attachment.Name}");
                    }
                }
                else
                {
                    Console.WriteLine($"Email {message.Id} has no attachments.");
                }
            }
        }
        private async Task DownloadAndUploadAttachments(string messageId, string userEmail, string sharePointSiteUrl, string targetFolder)
        {
            var messageContent = await _graphClient.Users[userEmail].Messages[messageId].Request().GetAsync();

            foreach (var attachment in messageContent.Attachments)
            {
                var attachmentBytes = await messageContent.Attachments[attachment.Id].Content.Request().ReadAsByteArrayAsync();
                await UploadAttachmentToSharePoint(attachment.Name, attachmentBytes, sharePointSiteUrl, targetFolder);
                Console.WriteLine($"Uploaded attachment: {attachment.Name}");
            }
        }

        private async Task UploadAttachmentToSharePoint(string fileName, byte[] attachmentBytes, string sharePointSiteUrl, string targetFolder)
        {
            string uploadUrl = $"{sharePointSiteUrl}/sites/root/drives/me/items/'{targetFolder}'/Files/add(overwrite=true, unconverted=true)";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await GetSharePointAccessTokenAsync());

                using (var content = new ByteArrayContent(attachmentBytes))
                {
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") { FileName = fileName };
                    var response = await client.PostAsync(uploadUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception("Failed to upload attachment to SharePoint");
                    }
                }
            }
        }

        public async Task<string> GetSharePointAccessTokenAsync()
        {
            return string.Empty;
        }
    }
}
