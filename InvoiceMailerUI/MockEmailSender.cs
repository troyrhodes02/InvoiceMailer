using System;
using System.IO;
using System.Threading.Tasks;
using Serilog;

namespace InvoiceMailerUI
{
    /// <summary>
    /// A mock implementation of IEmailSender that logs email information without actually sending emails.
    /// Useful for testing and development.
    /// </summary>
    public class MockEmailSender : IEmailSender
    {
        private string _userEmail = "mock@example.com";
        private string _overrideSenderEmail = string.Empty;

        public Task<bool> AuthenticateAsync(bool silent = false)
        {
            // Mock successful authentication
            Log.Information("MOCK EMAIL SENDER: Simulating successful authentication");
            Console.WriteLine("[TEST MODE] User would be prompted to log in");
            
            // In mock mode, we just pretend authentication was successful
            return Task.FromResult(true);
        }

        public string GetUserEmail()
        {
            return _userEmail;
        }

        public void SetSenderEmail(string senderEmail)
        {
            _overrideSenderEmail = senderEmail;
            Log.Information("MOCK EMAIL SENDER: Sender email override set to {SenderEmail}", senderEmail);
        }

        public string GetEffectiveSenderEmail()
        {
            return !string.IsNullOrEmpty(_overrideSenderEmail) ? _overrideSenderEmail : _userEmail;
        }

        public Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null)
        {
            // Determine effective sender
            string senderEmail = GetEffectiveSenderEmail();
            
            // Log details instead of sending a real email
            Log.Information("MOCK EMAIL SENDER: Would send email to {Recipient}", toEmail);
            Log.Information("MOCK EMAIL SENDER: Subject: {Subject}", subject);
            Log.Information("MOCK EMAIL SENDER: Body: {Body}", bodyText);
            Log.Information("MOCK EMAIL SENDER: From: {From}", senderEmail);
            
            if (!string.IsNullOrEmpty(attachmentPath))
            {
                if (File.Exists(attachmentPath))
                {
                    Log.Information("MOCK EMAIL SENDER: Would attach file: {FilePath}", attachmentPath);
                }
                else
                {
                    Log.Warning("MOCK EMAIL SENDER: Attachment file not found: {FilePath}", attachmentPath);
                }
            }
            
            Console.WriteLine($"[TEST MODE] Email would be sent to {toEmail} from {senderEmail}");
            
            // Just return a completed task since we're not actually sending anything
            return Task.CompletedTask;
        }
    }
} 