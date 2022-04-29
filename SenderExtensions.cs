using MailKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace MailKitExtensions.Sender
{
    public static class SenderExtensions
    {
        #region Initialisation
        private static ILogger? _initLog;
        private static ILogger _logger
        {
            get
            {
                if (_initLog == null)
                {
                    ILoggerProvider loggerProvider = new DebugLoggerProvider();
                    var loggerProviders = new ILoggerProvider[] { loggerProvider };
                    var loggerFactory = new LoggerFactory(loggerProviders);
                    _initLog = loggerFactory.CreateLogger(nameof(SenderExtensions));
                }
                return _initLog ?? NullLogger.Instance;
            }
        }
        #endregion

        #region File Handler
        public static bool FileCheckOk(
            string filePath, bool checkFile = false)
        {
            bool isExisting = false;

            // Directory must end with a "\\" character
            if (!checkFile &&
                !Path.HasExtension(filePath) &&
                !filePath.EndsWith(Path.DirectorySeparatorChar))
            {
                var sb = new StringBuilder(filePath);
                sb.Append(Path.DirectorySeparatorChar);
                filePath = sb.ToString();
            }

            var directory = Path.GetDirectoryName(filePath);

            if (!Directory.Exists(directory))
                _logger.LogInformation("Folder not found: '{0}'.", directory);
            else if (checkFile && !File.Exists(filePath))
                _logger.LogInformation("File not found: '{0}'.", filePath);
            else
                isExisting = true;

            return isExisting;
        }

        public static async Task<Stream> GetFileStreamAsync(
            string filePath, CancellationToken ct = default)
        {
            const int BufferSize = 8192;
            var outputStream = new MemoryStream();
            if (FileCheckOk(filePath, true))
            {
                using (var source = new FileStream(
                    filePath, FileMode.Open, FileAccess.Read,
                    FileShare.Read, BufferSize, useAsync: true))
                {
                    //var bytes = new byte[source.Length];
                    //await source.ReadAsync(bytes, 0, bytes.Length, ct);
                    //await outputStream.WriteAsync(bytes, 0, bytes.Length, ct);
                    await source.CopyToAsync(outputStream, ct);
                }
                outputStream.Position = 0;
                _logger.LogDebug("OK '{0}'.", filePath);
            }
            return outputStream;
        }
        #endregion

        #region Mime Entity
        public static MimeEntity? GetMimePart(
            this Stream stream, string fileName, string contentType = "", string contentId = "")
        {
            MimeEntity? result = null;
            if (stream != null && stream.Length > 0)
            {
                stream.Position = 0; // reset stream position ready to read
                if (string.IsNullOrWhiteSpace(contentType))
                    contentType = System.Net.Mime.MediaTypeNames.Application.Octet;
                if (string.IsNullOrWhiteSpace(contentId))
                    contentId = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                //streamIn.CopyTo(streamOut, 8192);
                result = new MimePart(contentType)
                {
                    Content = new MimeContent(stream),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentId = contentId,
                    FileName = fileName
                };
            }
            return result;
        }
        #endregion

        #region Mime Messages
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
            bool send = false;

            if (client == null)
                _logger.LogDebug("No client was supplied, message not sent.");
            else if (message == null)
                _logger.LogDebug("No message was supplied, none was sent.");
            else if (message.IsCircularReference())
                _logger.LogInformation("Circular reference detected, message not sent.");
            else
            {
                try
                {
                    await client.SendAsync(message, ct);
                    send = true;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to send email, MessageId='{0}'.", message.MessageId);
                }
            }

            return send;
        }

        public static async Task<bool> SendMimeMessageAsync(
            this SmtpClient client, IEnumerable<MimeMessage> messages, CancellationToken ct = default)
        {
            bool isSend = messages != null ? true : false;

            if (isSend)
            {
                foreach (var message in messages)
                {
                    isSend &= await client.SendMimeMessageAsync(message, ct);
                }
            }
            else
                _logger.LogDebug("No messages were supplied, none were sent.");

            return isSend;
        }

        public static bool IsCircularReference(this MimeMessage message)
        {
            if (message != null && message.From.Count == 0)
            {
                _logger.LogTrace("No 'From' address was set, adding '{0}'. Date: {1}. Subject: '{2}'.",
                    EmailSender.FromMailboxAddress.Address, message.Date, message.Subject);
                message.From.Add(EmailSender.FromMailboxAddress);
            }
            if (message != null && message.To.Count == 0 && message.Cc.Count == 0 && message.Bcc.Count == 0)
            {
                _logger.LogDebug("No 'To' address was set, adding '{0}'. Date: {1}. Subject: '{2}'.",
                    EmailSender.ToMailboxAddressesEnumerated, message.Date, message.Subject);
                message.To.AddRange(EmailSender.ToMailboxAddresses);
            }
            var mailboxToCcAddresses = message == null ? Array.Empty<MailboxAddress>() :
                Enumerable.Concat(message.To.Mailboxes, message.Cc.Mailboxes);
            return mailboxToCcAddresses.EmailAddressesIntersect(message?.From.Mailboxes);
        }

        public static bool EmailAddressesIntersect(
            this IEnumerable<MailboxAddress> mailboxes1, IEnumerable<MailboxAddress> mailboxes2)
        {
            return mailboxes1?.Select(m => m.Address)?.Intersect(mailboxes2
                ?.Select(m => m.Address), StringComparer.OrdinalIgnoreCase)?.Any() ?? false;
        }
        #endregion
    }
}
