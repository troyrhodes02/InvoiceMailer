using System.Threading.Tasks;

namespace InvoiceMailerUI
{
    public interface IEmailSender
    {
        Task<bool> AuthenticateAsync(bool silent = false);
        string GetUserEmail();
        void SetSenderEmail(string senderEmail);
        string GetEffectiveSenderEmail();
        Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null);
    }
} 