using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System.IO;
using System.Threading.Tasks;

namespace InvoiceMailerUI
{
    public class EmailSender : IEmailSender
    {
        private readonly GraphServiceClient _graphClient;
        private readonly string _senderEmail;

        public EmailSender(string tenantId, string clientId, string clientSecret, string senderEmail)
        {
            // Create the client credentials
            var clientSecretCredential = new ClientSecretCredential(
                tenantId,
                clientId,
                clientSecret);

            // Initialize the Graph client
            _graphClient = new GraphServiceClient(clientSecretCredential, 
                new[] { "https://graph.microsoft.com/.default" });
                
            // Store the sender email
            _senderEmail = senderEmail;
        }

        public async Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null)
        {
            // Create the message
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = bodyText
                },
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = toEmail
                        }
                    }
                },
                // Explicitly set the sender
                From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = _senderEmail
                    }
                }
            };

            // Add attachment if path is provided and file exists
            if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
            {
                var fileBytes = await File.ReadAllBytesAsync(attachmentPath);
                var fileName = Path.GetFileName(attachmentPath);

                message.Attachments = new List<Attachment>
                {
                    new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        ContentBytes = fileBytes,
                        ContentType = "application/pdf", // Assuming PDF, adjust if needed
                        Name = fileName
                    }
                };
            }

            // Create the message request
            var requestBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };

            try
            {
                // Try to send using the shared mailbox approach (if available)
                await _graphClient.Users[_senderEmail].SendMail.PostAsync(requestBody);
            }
            catch (ODataError ex) when (ex.Error?.Message?.Contains("Access is denied") == true)
            {
                // If access denied, try alternative approach
                // Note: This may require additional permissions in Azure AD
                // and might not work depending on your tenant configuration
                throw new InvalidOperationException(
                    "Email sending failed. The application doesn't have permission to send mail as this user. " +
                    "Please ask your Azure AD administrator to grant the application the necessary permissions.", ex);
            }
        }
    }
} 