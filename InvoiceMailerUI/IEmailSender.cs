using System.Threading.Tasks;

namespace InvoiceMailerUI
{
    public interface IEmailSender
    {
        Task SendEmailAsync(
            string toEmail,
            string subject,
            string bodyText,
            string? attachmentPath = null);
    }
} 