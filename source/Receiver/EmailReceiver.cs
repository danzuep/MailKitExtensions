using MimeKit;
using MailKit;
using MailKit.Net.Imap;
using System.Net;
using Microsoft.Extensions.Logging;
using Zue.Common;
using MailKitExtensions.Models;

namespace MailKitExtensions.Receiver
{
    public class EmailReceiver : IEmailReceiver
    {
        #region Mail Folders Properties
        private object _mailFolderLock = new object();
        private IMailFolder? _mailFolder;
        public IMailFolder MailFolder
        {
            get
            {
                Reconnect();
                lock (_mailFolderLock)
                    if (_mailFolder == null)
                        lock (_imapClient.SyncRoot)
                            _mailFolder = string.IsNullOrEmpty(_folderName) ?
                                _imapClient.Inbox : _imapClient.GetFolder(_folderName);
                return _mailFolder;
            }
            internal set
            {
                _mailFolder = value;
                _folderName = value.FullName;
            }
        }

        private object _sentFolderLock = new object();
        private IMailFolder? _sentFolder;
        public IMailFolder? SentFolder
        {
            get
            {
                Reconnect();
                lock (_sentFolderLock)
                    if (_sentFolder == null)
                        lock (_imapClient.SyncRoot)
                            _sentFolder = _imapClient.GetSentFolder();
                return _sentFolder;
            }
        }
        #endregion

        #region Public Properties
        public static MessageSummaryItems ItemFilter
        {
            get =>
                MessageSummaryItems.Envelope |
                MessageSummaryItems.BodyStructure |
#if DEBUG
                MessageSummaryItems.Size |
                //MessageSummaryItems.Flags |
                MessageSummaryItems.InternalDate |
#endif
                MessageSummaryItems.UniqueId;
        }
        #endregion

        #region Private and Internal Fields
        internal static ILogger _logger = LogUtil.GetLogger<EmailReceiver>();
        internal readonly ImapClient _imapClient;
        private const string _logFilePath = "C:\\Temp\\EmailClientImap.log";
        private readonly ProtocolLogger? _imapLogger;
        private readonly string _imapHost;
        //private readonly int _imapPort;
        private readonly NetworkCredential _credential;
        internal string _folderName = "INBOX";
        #endregion

        public EmailReceiver(EmailReceiverOptions receiverOptions, bool connect = true, bool useLogger = false)
        {
            _credential = receiverOptions.ImapCredential;
            _imapHost = receiverOptions.ImapHost;
            _folderName = receiverOptions.FolderToProcess;
            if (useLogger && _imapLogger is null)
                _imapLogger = GetProtocolLogger(_logFilePath);
            _imapClient = GetImapClient(_imapHost, _credential, connect, _imapLogger);
        }

        public EmailReceiver(NetworkCredential credential, string imapHost,
            string folderName = "INBOX", bool connect = true, bool useLogger = false)
        {
            if (string.IsNullOrWhiteSpace(imapHost))
                throw new ArgumentException(nameof(imapHost));
            _credential = credential;
            _imapHost = imapHost;
            _folderName = string.IsNullOrEmpty(folderName) ? "INBOX" : folderName;
            if (useLogger && _imapLogger is null)
                _imapLogger = GetProtocolLogger(_logFilePath);
            _imapClient = GetImapClient(_imapHost, _credential, connect, _imapLogger);
        }

        public static ImapClient GetImapClient(
            string imapHost, NetworkCredential credential, bool connect = true, ProtocolLogger? imapLogger = null)
        {
            var mailClient = imapLogger is null ? new ImapClient() : new ImapClient(imapLogger);
            if (connect)
                mailClient.Connect(imapHost, credential);
            return mailClient;
        }

        internal static ProtocolLogger? GetProtocolLogger(string? logFilePath = null)
            => string.IsNullOrWhiteSpace(logFilePath) ?
                new ProtocolLogger(Console.OpenStandardError()) :
                new ProtocolLogger(logFilePath, false);

        public void Reconnect(CancellationToken ct = default)
        {
            if (_imapClient is null)
                throw new InvalidOperationException(nameof(_imapClient));
            _imapClient.Connect(_imapHost, _credential, ct: ct);
        }

        public IMailFolder ConnectFolder(bool enableWrite = false, CancellationToken ct = default)
        {
            Reconnect(ct);
            return MailFolder.ConnectFolder(enableWrite, ct);
        }

        public IMailFolder OpenFolder(
            string folderName = "INBOX", bool openFolder = true, bool enableWrite = false, CancellationToken ct = default)
        {
            Reconnect(ct);

            if (_mailFolder == null || _mailFolder.FullName != folderName && !ct.IsCancellationRequested)
            {
                if (_mailFolder != null && _mailFolder.IsOpen)
                    //lock (_imapClient.SyncRoot)
                        lock (_mailFolder.SyncRoot)
                            _mailFolder.Close(false);
                _mailFolder = _imapClient.OpenFolder(folderName, openFolder, enableWrite, ct);
                _folderName = _mailFolder?.FullName ?? "INBOX";
            }

            return _mailFolder;
        }

        public async Task GetMailFoldersAsync(bool all = true, CancellationToken ct = default)
        {
            Reconnect(ct);
            await _imapClient.GetMailFoldersAsync(all, ct);
        }

        public async Task<int> MoveMessagesAsync(
            IEnumerable<UniqueId> messages, string destinationFolder,
            bool copy = false, bool markAsRead = true, CancellationToken ct = default)
        {
            return await _imapClient.MoveMessagesAsync(messages.ToList(), MailFolder, destinationFolder, copy, markAsRead, ct);
        }

        public async Task<(UniqueId?, IMailFolder?)> MoveToMailFolderAsync(
            UniqueId messageUid, IMailFolder source, string destinationFolder, bool copy = false,
            bool leaveUnread = false, string description = "", CancellationToken ct = default)
        {
            UniqueId? resultUid = null;
            IMailFolder? destination = null;
            if (!string.IsNullOrEmpty(destinationFolder) && messageUid.IsValid)
            {
                try
                {
                    Reconnect(ct);
                    var mailFolder = source ?? MailFolder;
                    source = mailFolder.ConnectFolder(true, ct);
                    lock (_imapClient.SyncRoot)
                        destination = _imapClient.GetFolder(destinationFolder, ct);
                    resultUid = await messageUid.MoveToMailFolderAsync(source, destination, copy, leaveUnread, description, ct);
                }
                catch (FolderNotFoundException ex)
                {
                    _logger.LogWarning(ex, "'{0}' folder not found, {1} not moved.", destinationFolder, messageUid);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "{0} {1} not moved.", source.FullName, messageUid);
                }
            }
            return (resultUid, destination);
        }

        public void AppendToSentFolder(IEnumerable<MimeMessage> messages, CancellationToken ct = default)
        {
            if (messages != null)
            {
                try
                {
                    foreach (var message in messages)
                        SentFolder?.AppendMessage(message, ct: ct);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to add {0} messages to the 'Sent Items' mail folder.", messages?.Count() ?? 0);
                }
            }
        }

        public void AppendToSentFolder(MimeMessage message, CancellationToken ct = default)
        {
            try
            {
                SentFolder?.AppendMessage(message, ct: ct);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to add message to the 'Sent Items' mail folder.");
            }
        }

        public virtual void Disconnect()
        {
            if (_imapClient?.IsConnected ?? false)
            {
                //if (_mailFolder?.IsOpen ?? false) _mailFolder?.Close(false); // Beware, throws an IOException
                //if (_mailFolder != null) _mailFolder = null;
                lock (_imapClient.SyncRoot)
                    _imapClient?.Disconnect(true);
            }
        }

        public virtual void Dispose()
        {
            //_logger.LogDebug("Disposing of the IMAP email receiver client...");
            Disconnect();
            _imapClient?.Dispose();
        }
    }

    public interface IEmailReceiver : IDisposable
    {
        void Reconnect(CancellationToken ct);
        IMailFolder OpenFolder(string folderName, bool openFolder, bool enableWrite, CancellationToken ct);
    }
}
