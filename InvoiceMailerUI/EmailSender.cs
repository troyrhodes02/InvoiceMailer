using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace InvoiceMailerUI
{
    public class EmailSender : IEmailSender
    {
        private GraphServiceClient? _graphClient;
        private string _userEmail = string.Empty;
        private string _overrideSenderEmail = string.Empty;
        private readonly string _tenantId;
        private readonly string _clientId;
        private readonly string[] _scopes = new string[] { "Mail.Send", "Mail.ReadWrite", "User.Read" };
        private IPublicClientApplication? _msalClient;

        public EmailSender(string tenantId, string clientId)
        {
            _tenantId = tenantId;
            _clientId = clientId;
        }

        /// <summary>
        /// Set an override sender email address
        /// </summary>
        public void SetSenderEmail(string senderEmail)
        {
            _overrideSenderEmail = senderEmail;
        }

        /// <summary>
        /// Get the effective sender email (override if set, otherwise authenticated user email)
        /// </summary>
        public string GetEffectiveSenderEmail()
        {
            return !string.IsNullOrEmpty(_overrideSenderEmail) ? _overrideSenderEmail : _userEmail;
        }

        /// <summary>
        /// Authenticate the user interactively using MSAL.NET and Azure.Identity
        /// </summary>
        public async Task<bool> AuthenticateAsync(bool silent = false)
        {
            try
            {
                // Initialize MSAL client application
                _msalClient = PublicClientApplicationBuilder
                    .Create(_clientId)
                    .WithAuthority(AzureCloudInstance.AzurePublic, _tenantId)
                    .WithRedirectUri("http://localhost")
                    .Build();

                // Create the interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = _tenantId,
                    ClientId = _clientId,
                    RedirectUri = new Uri("http://localhost")
                };

                TokenCredential credential;
                
                if (silent)
                {
                    try
                    {
                        // Try silent authentication first
                        var accounts = await _msalClient.GetAccountsAsync();
                        var firstAccount = accounts.FirstOrDefault();
                        
                        if (firstAccount != null)
                        {
                            // Use DefaultAzureCredential which attempts silent auth methods
                            var defaultOptions = new DefaultAzureCredentialOptions
                            {
                                TenantId = _tenantId,
                                ExcludeInteractiveBrowserCredential = true
                            };
                            credential = new DefaultAzureCredential(defaultOptions);
                            
                            // Validate the credential can get a token
                            var tokenRequestContext = new TokenRequestContext(_scopes);
                            await credential.GetTokenAsync(tokenRequestContext, CancellationToken.None);
                        }
                        else
                        {
                            // No cached account, return failure for silent auth
                            return false;
                        }
                    }
                    catch
                    {
                        // If silent auth is required but fails, return failure
                        if (silent)
                        {
                            return false;
                        }
                        
                        // Otherwise fall back to interactive
                        credential = new InteractiveBrowserCredential(options);
                    }
                }
                else
                {
                    // Use interactive browser when silent auth is not required
                    credential = new InteractiveBrowserCredential(options);
                }

                // Create a Graph client with the credential
                _graphClient = new GraphServiceClient(credential, _scopes);

                // Get user profile information to set email
                var user = await _graphClient.Me.GetAsync();
                if (user?.Mail != null)
                {
                    _userEmail = user.Mail;
                }
                else if (user?.UserPrincipalName != null)
                {
                    _userEmail = user.UserPrincipalName;
                }
                else
                {
                    throw new InvalidOperationException("Unable to retrieve user email");
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Authentication failed: {ex.Message}");
                return false;
            }
        }

        public string GetUserEmail()
        {
            return _userEmail;
        }

        public async Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null)
        {
            // Ensure we're authenticated first
            if (_graphClient == null)
            {
                throw new InvalidOperationException("Not authenticated. Call AuthenticateAsync before sending emails.");
            }

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
                }
                // From address is set automatically to the authenticated user
            };

            // If we have an override sender, set it explicitly
            if (!string.IsNullOrEmpty(_overrideSenderEmail))
            {
                message.From = new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = _overrideSenderEmail
                    }
                };
            }

            // Add attachment if path is provided and file exists
            if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
            {
                var fileBytes = await File.ReadAllBytesAsync(attachmentPath);
                var fileName = Path.GetFileName(attachmentPath);
                var contentType = GetContentType(fileName);

                message.Attachments = new List<Attachment>
                {
                    new FileAttachment
                    {
                        OdataType = "#microsoft.graph.fileAttachment",
                        ContentBytes = fileBytes,
                        ContentType = contentType,
                        Name = fileName
                    }
                };
            }

            // Create the message request
            var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true
            };

            try
            {
                // Send mail using the authenticated user's context
                await _graphClient.Me.SendMail.PostAsync(requestBody);
            }
            catch (ODataError ex)
            {
                throw new InvalidOperationException(
                    $"Email sending failed: {ex.Error?.Message}. " +
                    "Make sure your account has the necessary permissions.", ex);
            }
        }

        // Helper method to determine the content type based on file extension
        private string GetContentType(string fileName)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            
            return extension switch
            {
                ".pdf" => "application/pdf",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                _ => "application/octet-stream"
            };
        }
    }
} 