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
        public Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null)
        {
            // Log details instead of sending a real email
            Log.Information("MOCK EMAIL SENDER: Would send email to {Recipient}", toEmail);
            Log.Information("MOCK EMAIL SENDER: Subject: {Subject}", subject);
            Log.Information("MOCK EMAIL SENDER: Body: {Body}", bodyText);
            
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
            
            Console.WriteLine($"[TEST MODE] Email would be sent to {toEmail}");
            
            // Just return a completed task since we're not actually sending anything
            return Task.CompletedTask;
        }
    }
} 