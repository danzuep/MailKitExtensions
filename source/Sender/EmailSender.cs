using MimeKit;
using MailKit;
using MailKit.Net.Smtp;
using System.Net;
using Zue.Common;
using MailKitExtensions.Writer;

namespace MailKitExtensions.Sender
{
    public interface IEmailSender : IDisposable
    {
        Task<bool> SendAsync(string from, string to, string subject, string body, bool isHtml = true, CancellationToken ct = default, params string[] attachments);
    }

    public class EmailSender : IEmailSender
    {
        internal readonly ILog _logger = LogUtil.GetLogger<EmailSender>();

        public static MailboxAddress? DefaultFromAddress { get; set; }
        public static MailboxAddress? DefaultToAddress { get; set; }

        #region Private Fields
        private readonly string _smtpHost;
        private SmtpClient? _smtpClient;
        private ProtocolLogger? _smtpLogger;
        private ICredentials? _credential;
        #endregion

        public EmailSender(string smtpHost, string fromAddress = "",
            string toAddress = "", ICredentials? credential = null)
        {
            if (string.IsNullOrWhiteSpace(smtpHost))
                throw new ArgumentException(nameof(smtpHost));
            _smtpHost = smtpHost;
            _credential = credential;
            if (string.IsNullOrWhiteSpace(fromAddress))
                DefaultFromAddress = MailboxAddress.Parse(fromAddress);
            if (string.IsNullOrWhiteSpace(fromAddress))
                DefaultToAddress = MailboxAddress.Parse(toAddress);
        }

        public async Task<bool> SendAsync(
            string from, string to, string subject, string bodyText,
            bool isHtml = true, CancellationToken ct = default, params string[] attachmentNames)
        {
            var mimeMessage = EmailHelper.CreateMimeMessage(from, to, subject, bodyText, isHtml, "", attachmentNames);
            return await SendAsync(mimeMessage, ct);
        }

        public async Task<bool> SendAsync(MimeMessage email, CancellationToken ct = default)
            => await email.SendMimeMessageAsync(_smtpHost, _credential, ct);

        public async Task<bool> SendAsync(IEnumerable<MimeMessage> emails, CancellationToken ct = default)
            => await emails.SendMimeMessagesAsync(_smtpHost, _credential, ct);

        public async Task ConnectSmtpClientAsync(ICredentials? credential = null, bool useLogger = false)
        {
            _credential = credential;
            if (useLogger && _smtpLogger == null)
                _smtpLogger = new ProtocolLogger(Console.OpenStandardError());
            _smtpClient = await credential.ConnectSmtpClientAsync(_smtpHost, _smtpLogger);
        }

        public void DisposeSmtpClient()
        {
            if (_smtpClient?.IsConnected ?? false)
                _smtpClient?.Disconnect(true);
            _credential = null;
        }

        public void Dispose()
        {
            DisposeSmtpClient();
            _smtpLogger?.Dispose();
        }
    }
}