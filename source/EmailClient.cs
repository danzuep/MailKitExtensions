using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MimeKit;
using Zue.Common;
using MailKitExtensions.Models;
using MailKitExtensions.Sender;
using MailKitExtensions.Receiver;
using MailKitExtensions.Writer;

namespace MailKitExtensions
{
    public class EmailClient : IDisposable
    {
        private readonly ILogger _logger = LogProvider.GetLogger<EmailClient>();
        private readonly EmailOptions _emailOptions;
        public EmailSender Sender { get; private set; }
        public EmailReceiver Receiver { get; private set; }

        public EmailClient(IOptions<EmailOptions> emailOptions)
        {
            if (emailOptions?.Value == null)
                throw new ArgumentNullException(nameof(emailOptions));
            _emailOptions = emailOptions.Value;
            Sender = new EmailSender(_emailOptions.Sender.SmtpHost, _emailOptions.Sender.FromAddress, _emailOptions.Sender.ToAddress, _emailOptions.Sender.SmtpCredential);
            Receiver = new EmailReceiver(_emailOptions.Receiver.ImapCredential, _emailOptions.Receiver.ImapHost, _emailOptions.Receiver.FolderToProcess, false);
        }

        public async Task<bool> SendAsync(MimeMessage mimeMessage, bool appendSent = true, CancellationToken ct = default)
        {
            bool isSend = mimeMessage != null;
            if (isSend)
            {
                isSend &= await Sender.SendAsync(mimeMessage!, ct);
                if (appendSent && isSend)
                    Receiver.AppendToSentFolder(mimeMessage!);
            }
            return isSend;
        }

        public async Task<bool> SendAsync(IEnumerable<MimeMessage> emails, bool appendSent = true, CancellationToken ct = default)
        {
            bool isSend = emails != null;
            if (isSend)
            {
                using var client = await SenderExtensions.GetSmtpClientAsync(
                    _emailOptions.Sender.SmtpHost, _emailOptions.Sender.SmtpCredential, ct);
                foreach (var email in emails!)
                {
                    isSend &= await client.SendMimeMessageAsync(email, ct);
                    if (appendSent && isSend)
                        Receiver.AppendToSentFolder(email);
                }
                await client.DisconnectAsync(true, ct);
            }
            return isSend;
        }

        public async Task<bool> SendEmailAsync(
            string to, string subject, string? body = null, string? from = null, bool isHtml = true, bool appendSent = true, CancellationToken ct = default, params string[] fileNames)
        {
            if (string.IsNullOrEmpty(from))
                from = _emailOptions.Sender.FromAddress;
            if (string.IsNullOrEmpty(to))
                to = _emailOptions.Sender.ToAddress;
            if (string.IsNullOrEmpty(subject))
                subject = _emailOptions.Sender.Subject;
            if (string.IsNullOrEmpty(body))
                body = _emailOptions.Sender.BodyHtml.Replace("{{BodyText}}", _emailOptions.Sender.BodyText);
            else if (isHtml)
                body = _emailOptions.Sender.BodyHtml.Replace("{{BodyText}}", body);

            body = body.Replace("{{Signature}}", _emailOptions.Sender.Signature
                .Replace("{{DateTime.Now}}", DateTimeOffset.Now.ToString(IdleClientReceiver.DateTimeFormatLong))); //G
            var mimeMessage = EmailHelper.CreateMimeMessage(from, to, subject, body, isHtml, "", fileNames);

            return await SendAsync(mimeMessage, appendSent, ct);
        }

        public async Task EmailStackTrace<T>(
            Exception? ex, string whatHappened = "", LogLevel logLevel = LogLevel.Error) where T : class
        {
            ex = ex?.InnerException ?? ex;
            if (string.IsNullOrWhiteSpace(whatHappened))
                whatHappened = ex?.Message ?? "";
            _logger.Log(logLevel, ex, whatHappened);
            string subject = string.Join(' ', nameof(T), nameof(logLevel));
            string description = string.Join(": ", subject, whatHappened);
            string body = whatHappened + "<br /><br />" + ex?.ToString();
            await SendEmailAsync("", subject, body, appendSent: false);
        }

        public void Dispose()
        {
            Sender?.Dispose();
            Receiver?.Dispose();
        }
    }
}
