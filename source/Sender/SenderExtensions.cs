using MimeKit;
using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using System.Net;
using Zue.Common;

namespace MailKitExtensions.Sender
{
    public static class SenderExtensions
    {
        private static ILog _logger = LogUtil.GetLogger(nameof(SenderExtensions));

        public static async Task<SmtpClient> GetSmtpClientAsync(
            string host, ICredentials? credential = null, CancellationToken ct = default, ProtocolLogger? smtpLogger = null)
        {
            var mailClient = smtpLogger == null ? new SmtpClient() : new SmtpClient(smtpLogger);
            try
            {
                if (!mailClient.IsConnected && !string.IsNullOrEmpty(host))
                    await mailClient.ConnectAsync(host, cancellationToken: ct);
                if (credential != null && !mailClient.IsAuthenticated)
                    await mailClient.AuthenticateAsync(credential, ct);
            }
            catch (AuthenticationException ex)
            {
                _logger.LogDebug("{0}: Failed to authenticate.", ex.GetType().Name);
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogDebug("{0}: Failed to connect to email client.", ex.GetType().Name);
                throw;
            }
            return mailClient;
        }

        internal static async Task<SmtpClient> ConnectSmtpClientAsync(
            this ICredentials? credential, string smtpHost, ProtocolLogger? smtpLogger = null, CancellationToken ct = default)
        {
            var client = await GetSmtpClientAsync(smtpHost, credential, ct, smtpLogger);
            client.Connect(smtpHost, credential, ct: ct);
            return client;
        }

        public static void Connect(this SmtpClient client, string host,
            ICredentials? credential = null, CancellationToken ct = default)
        {
            if (client == null)
                throw new ArgumentNullException(nameof(client));

            if (string.IsNullOrEmpty(host))
                throw new ArgumentNullException(nameof(host));

            bool useSsl = credential == null ? false : true;
            // SMTP port 25 or 587 unencrypted, 465 encrypted
            int port = useSsl ? 465 : 25;

            if (!client.IsConnected)
            {
                client.Connect(host, port, useSsl, ct);
            }

            if (useSsl && !client.IsAuthenticated)
            {
                client.Authenticate(credential, ct);
            }
        }

        public static async Task<bool> SendMimeMessageAsync(
            this MimeMessage message, string smtpHost, ICredentials? credential = null, CancellationToken ct = default)
        {
            using var client = await GetSmtpClientAsync(smtpHost, credential, ct);
            bool send = await client.SendMimeMessageAsync(message, ct);
            await client.DisconnectAsync(true, ct);
            return send;
        }

        public static async Task<bool> SendMimeMessagesAsync(
            this IEnumerable<MimeMessage> mimeMessages, string smtpHost, ICredentials? credential = null, CancellationToken ct = default)
        {
            bool isSend = mimeMessages != null ? true : false;

            if (isSend)
            {
                using var client = await GetSmtpClientAsync(smtpHost, credential, ct);
                isSend = await client.SendMimeMessageAsync(mimeMessages!, ct);
                await client.DisconnectAsync(true, ct);
            }
            else
                _logger.LogDebug("No messages were supplied, none were sent.");

            return isSend;
        }

        public static async Task<bool> SendMimeMessageAsync(
            this SmtpClient client, MimeMessage message, CancellationToken ct = default)
        {
            bool isSent = false;

            if (client == null)
                _logger.LogDebug("No client was supplied, message not sent.");
            else if (message == null)
                _logger.LogDebug("No message was supplied, none was sent.");
            else if (message.IsCircularReference())
                _logger.LogInformation("Circular reference detected, message not sent.");
            else
            {
                if (message.From.Count == 0)
                {
                    _logger.LogTrace("No 'From' address set.");
                    if (EmailSender.DefaultFromAddress != null)
                        message.From.Add(EmailSender.DefaultFromAddress);
                }
                if (message.To.Count == 0 && message.Cc.Count == 0 && message.Bcc.Count == 0)
                {
                    _logger.LogTrace("No 'To' addresses set.");
                    if (EmailSender.DefaultToAddress != null)
                        message.To.Add(EmailSender.DefaultToAddress);
                }
                try
                {
                    await client.SendAsync(message, ct);
                    isSent = true;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to send email, MessageId='{0}'.", message.MessageId);
                }
            }

            return isSent;
        }

        public static async Task<bool> SendMimeMessageAsync(
            this SmtpClient client, IEnumerable<MimeMessage> messages, CancellationToken ct = default)
        {
            int sentCount = 0;
            int messageCount = messages?.Count() ?? 0;

            if (messages != null && messageCount > 0)
                foreach (var message in messages)
                    if (await client.SendMimeMessageAsync(message, ct))
                        sentCount++;
            else
                _logger.LogDebug("No messages were supplied, none were sent.");

            return sentCount == messageCount;
        }

        public static bool IsCircularReference(this MimeMessage message)
        {
            var mailboxToCcAddresses = message == null ? Array.Empty<MailboxAddress>() :
                message.To.Mailboxes.Concat(message.Cc.Mailboxes);
            return mailboxToCcAddresses.EmailAddressesIntersect(message?.From.Mailboxes);
        }

        public static bool EmailAddressesIntersect(
            this IEnumerable<MailboxAddress> mailboxes1, IEnumerable<MailboxAddress> mailboxes2)
        {
            bool isIntersect = false;
            if (mailboxes1 != null && mailboxes2 != null)
                isIntersect = mailboxes1.Select(m => m.Address).Intersect(mailboxes2
                    .Select(m => m.Address), StringComparer.OrdinalIgnoreCase).Any();
            return isIntersect;
        }
    }
}
