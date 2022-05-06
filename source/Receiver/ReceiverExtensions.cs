using MimeKit;
using MimeKit.Tnef;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using Zue.Common;
using MailKitExtensions.Sender;

namespace MailKitExtensions.Receiver
{
    public static class ReceiverExtensions
    {
        private static ILogger _logger = LogUtil.GetLogger(nameof(ReceiverExtensions));

        #region Mime Message Body
        public static string GetBodyPlainText(this MimeMessage mimeMessage)
        {
            string result = "";
            if (mimeMessage != null)
            {
                bool isHtml = mimeMessage.HtmlBody != null;
                var body = isHtml ? mimeMessage.HtmlBody : mimeMessage.TextBody;
                if (body != null)
                    result = GetPlainTextBody(body, isHtml);
            }
            return result;
        }

        public static string GetPlainTextBody(string mailBody, bool isHtml = true)
        {
            string bodyText = string.Empty;
            string fromPattern = "\nFrom:";
            string subjectPattern = "\nSubject:";
            if (mailBody == null)
                mailBody = string.Empty;

            if (isHtml)
            {
                bodyText = ReceiverMimeMessages.DecodeHtmlBody(mailBody);
                if (bodyText == null)
                {
                    var matches = Regex.Matches(mailBody, @">(.*?)<", RegexOptions.Singleline);
                    var filtered = matches.Select(m => m.Groups[1].Value.Replace("\r", "").Replace("\n", ""))
                        .Where(m => !string.IsNullOrWhiteSpace(m));
                    bodyText = matches.Count > 0 ? filtered.ToEnumeratedString("\r\n") : "";
                }
            }
            else
            {
                bodyText = mailBody;
            }

            var bodyParts = Regex.Split(bodyText, fromPattern, RegexOptions.IgnoreCase);
            for (int i = 0; i < bodyParts.Length; i++)
            {
                var mail = Regex.Split(bodyParts[i], subjectPattern, RegexOptions.IgnoreCase);
                bodyParts[i] = mail.Length > 1 ? mail[1] : bodyParts[i];
            }
            if (bodyParts.Length > 1)
            {
                string delimiter = "\r\n" + subjectPattern;
                bodyText = string.Join(delimiter,
                    bodyParts[0], bodyParts[1]);
            }
            else if (bodyParts.Length > 0)
            {
                bodyText = bodyParts[0];
            }

            return bodyText;
        }
        #endregion

        #region IMessageSummary Body
        public static async Task<string> GetMailBodyTextAsync(
            this IMessageSummary mail, CancellationToken ct = default)
        {
            var textEntity = await mail.GetMailBodyEntityAsync(ct);
            return textEntity.GetBodyFullText();
        }

        public static string GetBodyFullText(this MimeEntity body) =>
            body is TextPart tp ? tp.Text ?? "" : "";

        public static async Task<MimeEntity> GetMailBodyEntityAsync(
            this IMessageSummary mail, CancellationToken ct = default)
        {
            MimeEntity result = null;
            if (!mail?.Folder?.IsOpen ?? false)
                await mail.Folder.OpenAsync(FolderAccess.ReadOnly, ct);
            if (mail?.HtmlBody != null)
            {
                // this will download *just* the text/html part
                result = await mail.Folder.GetBodyPartAsync(mail.UniqueId, mail.HtmlBody, ct);

            }
            else if (mail?.TextBody != null)
            {
                // this will download *just* the text/plain part
                result = await mail.Folder.GetBodyPartAsync(mail.UniqueId, mail.TextBody, ct);
            }
            return result;
        }

        public static IEnumerable<BodyPartBasic> GetMailAttachmentDetails(
            this IMessageSummary message, IList<string> suffix, bool getAllNonText = false)
        {
            IEnumerable<BodyPartBasic> attachments = Array.Empty<BodyPartBasic>();
            if (message?.Body is BodyPartMultipart multipart)
            {
                var attachmentParts = multipart.BodyParts.OfType<BodyPartBasic>()
                    .Where(p => (getAllNonText && !p.ContentType.MediaType
                        .Contains("text", StringComparison.OrdinalIgnoreCase)) || p.IsAttachment);
                attachments = suffix?.Count > 0 ? attachmentParts.Where(a => suffix.Any(s => a.FileName?
                    .EndsWith(s, StringComparison.OrdinalIgnoreCase) ?? false)) : attachmentParts;
            }
            return attachments.ToList();
        }
        #endregion

        public static async Task GetMailFoldersAsync(
            this ImapClient client, bool all = true, CancellationToken ct = default)
        {
            if (client == null)
                throw new ArgumentNullException(nameof(client));
            IList<IMailFolder> folders;
            lock(client.SyncRoot)
                folders = client.PersonalNamespaces?.Count > 0 ?
                    client.GetFolders(client.PersonalNamespaces[0], cancellationToken: ct) :
                    client.Inbox.GetSubfolders(cancellationToken: ct);
            await folders.GetMailboxInfoAsync(all);
        }

        public static async Task GetMailboxInfoAsync(
            this IList<IMailFolder> mailFolders, bool all = false, CancellationToken ct = default)
        {
            if (mailFolders?.Count > 0)
            {
                foreach (var folder in mailFolders)
                {
                    await folder.OpenAsync(FolderAccess.ReadOnly, ct);
                    if (all || folder.Count > 0)
                        _logger.LogDebug("[folder] {0} ({1})",
                            folder.FullName, folder.Count);
                    await folder.CloseAsync(false, ct);
                }
            }
            else
                _logger.LogInformation("No IMailFolder folders.");
        }

        public static async Task<IList<IMessageSummary>> GetMessageSummariesAsync(
            this IMailFolder mailFolder, MessageSummaryItems filter = MessageSummaryItems.UniqueId, CancellationToken ct = default)
        {
            IList<IMessageSummary> messageSummaries = new List<IMessageSummary>();
            if (mailFolder != null)
            {
                bool closeWhenFinished = !mailFolder.IsOpen;
                if (closeWhenFinished)
                    await mailFolder.OpenAsync(FolderAccess.ReadOnly, ct);
                _logger.LogDebug("[folder] {0} ({1})", mailFolder.FullName, mailFolder.Count);
                messageSummaries = await mailFolder.FetchAsync(0, -1, filter, ct);
                if (closeWhenFinished)
                    await mailFolder.CloseAsync(false, ct);
            }
            return messageSummaries;
        }

        public static IMailFolder? GetSentFolder(this ImapClient client, CancellationToken ct = default)
        {
            if (client == null)
                throw new ArgumentNullException(nameof(client));
            IMailFolder? sentFolder = null;
            if ((client.Capabilities & (ImapCapabilities.SpecialUse | ImapCapabilities.XList)) != 0)
            {
                lock(client.SyncRoot)
                    sentFolder = client.GetFolder(SpecialFolder.Sent);
            }
            else
            {
                string[] commonSentFolderNames = { "Sent Items", "Sent Mail", "Sent Messages" };
                lock (client.SyncRoot)
                    sentFolder = client.GetFolder(client.PersonalNamespaces[0]);
                lock (sentFolder.SyncRoot)
                    sentFolder = sentFolder.GetSubfolders(false, ct).FirstOrDefault(x =>
                        commonSentFolderNames.Contains(x.Name, StringComparer.OrdinalIgnoreCase));
            }
            return sentFolder;
        }

        public static void AppendMessage(this IMailFolder sentFolder, MimeMessage message, bool isSeen = true, CancellationToken ct = default)
        {
            if (sentFolder != null && message != null)
            {
                if (message.From.Count == 0 && EmailSender.DefaultFromAddress != null)
                    message.From.Add(EmailSender.DefaultFromAddress);

                var messageFlag = isSeen ? MessageFlags.Seen : MessageFlags.None;
                lock (sentFolder.SyncRoot)
                    sentFolder.Append(message, messageFlag, ct);
            }
        }

        public static async Task<int> CopyMessagesAsync(this IList<UniqueId> messageUids,
            IMailFolder source, IMailFolder dest, bool markAsRead = true, CancellationToken ct = default)
        {
            int numberChanged = messageUids?.Count ?? 0;
            if (numberChanged > 0 && source != null && dest != null)
            {
                string logMessage = string.Format("{0} messages copied from '{1}' to '{2}': {{0}}",
                    numberChanged, source.FullName, dest.FullName);
                if (!source.IsOpen)
                    lock (source.SyncRoot)
                        source.Open(FolderAccess.ReadOnly, ct);
                var copied = await source.CopyToAsync(messageUids, dest, ct);
                _logger.LogDebug(logMessage, copied.ToEnumeratedString());
                if (markAsRead)
                {
                    lock (dest.SyncRoot)
                    {
                        if (!dest.IsOpen || dest.Access != FolderAccess.ReadWrite)
                            dest.Open(FolderAccess.ReadWrite, ct);
                        var destUids = copied.Select(c => c.Value).ToList();
                        dest.AddFlags(destUids, MessageFlags.Seen, true, ct);
                        dest.Close(false, ct);
                    }
                }
            }
            return numberChanged;
        }

        public static async Task<int> MoveMessagesAsync(this IList<UniqueId> messageUids,
            IMailFolder source, IMailFolder dest, bool markAsRead = true, CancellationToken ct = default)
        {
            int numberChanged = messageUids?.Count ?? 0;
            if (numberChanged > 0 && source != null && dest != null)
            {
                string logMessage = string.Format("{0} messages {{0}} from '{1}' to '{2}': {{1}}",
                    numberChanged, source.FullName, dest.FullName);
                try
                {
                    if (!source.IsOpen || source.Access != FolderAccess.ReadWrite)
                        lock (source.SyncRoot)
                            source.Open(FolderAccess.ReadWrite, ct);
                    if (markAsRead)
                        lock (source.SyncRoot)
                            source.AddFlags(messageUids, MessageFlags.Seen, true, ct);
                    var moved = await source.MoveToAsync(messageUids, dest, ct);
                    _logger.LogDebug(logMessage, numberChanged, "moved",
                        source.FullName, dest.FullName, moved.ToEnumeratedString());
                }
                catch (InvalidOperationException ex) //covers ServiceNotConnectedException
                {
                    _logger.LogWarning(ex, ex.Message);
                    lock (source.SyncRoot)
                    {
                        var moved = source.MoveTo(messageUids, dest, ct);
                        _logger.LogInformation(logMessage, numberChanged, "moved without changing 'seen' status",
                            source.FullName, dest.FullName, moved.ToEnumeratedString());
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to move from '{0}' to '{1}'.", source.FullName, dest.FullName);
                }
            }
            return numberChanged;
        }

        public static async Task<UniqueId?> MoveToMailFolderAsync(this UniqueId messageUid, IMailFolder source,
            IMailFolder destination, bool markAsRead = true, string description = "", CancellationToken ct = default)
        {
            UniqueId? resultUid = null;
            if (source != null && destination != null)
            {
                string logMessage = string.Format("'{0}' {1} {{0}} to {{1}} '{2}'. {3}",
                    source.FullName, messageUid, destination.FullName, description);
                try
                {
                    //source.ConnectFolder(true, ct);
                    lock (source.SyncRoot)
                    {
                        if (!source.IsOpen || source.Access != FolderAccess.ReadWrite)
                        {
                            _logger.LogWarning("Move requires ReadWrite access for '{0}' mail folder, " +
                                "this will change the shared SyncRoot.", source.FullName);
                            source.Open(FolderAccess.ReadWrite, ct);
                        }
                        if (markAsRead)
                            source.FlagMessage(messageUid, MessageFlags.Seen, true, ct);
                        resultUid = source.MoveTo(messageUid, destination, ct);
                    }
                    if (resultUid == null)
                        _logger.LogWarning(logMessage, "failed to move", resultUid);
                    else
                        _logger.LogTrace(logMessage, "moved", resultUid);
                }
                catch (InvalidOperationException ex) //covers ServiceNotConnectedException
                {
                    _logger.LogWarning(ex, ex.Message);
                    resultUid = await source.MoveToAsync(messageUid, destination, ct);
                    _logger.LogInformation(logMessage, "moved without changing 'seen' status", resultUid);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to move '{0}' {1} to '{2}'.", source.FullName, messageUid, destination.FullName);
                }
            }
            return resultUid;
        }

        public static async Task<UniqueId?> CopyToMailFolderAsync(this UniqueId messageUid, IMailFolder source,
            IMailFolder destination, bool markAsRead = true, string description = "", CancellationToken ct = default)
        {
            UniqueId? resultUid = null;
            if (source != null && destination != null && messageUid.IsValid)
            {
                if (!source.IsOpen)
                    await source.OpenAsync(FolderAccess.ReadOnly, ct);
                resultUid = await source.CopyToAsync(messageUid, destination, ct);
                _logger.LogTrace("'{0}' {1} copied to {2} in '{3}'. {4}",
                    source.FullName, messageUid, resultUid, destination.FullName, description);
                if (markAsRead && resultUid != null)
                {
                    lock (destination.SyncRoot)
                    {
                        if (!destination.IsOpen || destination.Access != FolderAccess.ReadWrite)
                            destination.Open(FolderAccess.ReadWrite, ct);
                        destination.AddFlags(resultUid.Value, MessageFlags.Seen, true, ct);
                        destination.Close(false, ct);
                    }
                }
            }
            return resultUid;
        }

        public static async Task<UniqueId?> MoveToMailFolderAsync(this UniqueId messageUid, IMailFolder source,
            IMailFolder destination, bool copy = false, bool leaveUnread = false, string description = "", CancellationToken ct = default)
        {
            UniqueId? resultUid = null;
            if (source != null && destination != null && messageUid.IsValid)
            {
                try
                {
                    if (copy)
                        resultUid = await messageUid.CopyToMailFolderAsync(source, destination, !leaveUnread, description, ct);
                    else
                        resultUid = await messageUid.MoveToMailFolderAsync(source, destination, !leaveUnread, description, ct);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "{0} {1} not moved.", source.FullName, messageUid);
                }
            }
            return resultUid;
        }

        public static async Task<int> MoveMessagesAsync(this ImapClient client, IList<UniqueId> messageUids,
            IMailFolder source, string destinationFolder, bool copy = false, bool markAsRead = true, CancellationToken ct = default)
        {
            int moveCount = 0;
            try
            {
                IMailFolder destination;
                lock (client.SyncRoot)
                    destination = client.GetFolder(destinationFolder, ct);
                if (copy)
                    moveCount = await messageUids.CopyMessagesAsync(source, destination, markAsRead);
                else
                    moveCount = await messageUids.MoveMessagesAsync(source, destination, markAsRead);
            }
            catch (FolderNotFoundException ex)
            {
                _logger.LogWarning(ex, "'{0}' folder not found, messages not moved.", destinationFolder);
            }
            return moveCount;
        }

        public static void FlagMessage(
            this IMailFolder mailFolder, UniqueId uniqueId, MessageFlags flag = MessageFlags.Seen, bool addFlag = true, CancellationToken ct = default)
        {
            if (mailFolder == null)
                throw new ArgumentNullException(nameof(mailFolder));

            //mailFolder.ConnectFolder(false, ct);
            lock (mailFolder.SyncRoot)
            {
                bool closeWhenFinished = !mailFolder.IsOpen;
                if (closeWhenFinished || mailFolder.Access != FolderAccess.ReadWrite)
                {
                    _logger.LogWarning("Flag requires ReadWrite access for '{0}' mail folder, " +
                        "this will change the shared SyncRoot.", mailFolder.FullName);
                    mailFolder.Open(FolderAccess.ReadWrite, ct);
                }
                if (addFlag)
                    mailFolder.AddFlags(uniqueId, flag, true, ct);
                else
                    mailFolder.RemoveFlags(uniqueId, flag, true, ct);
                if (closeWhenFinished)
                    mailFolder.Close(false, ct);
            }

            _logger.LogTrace("'{Folder}' {ID} message has been flagged as {state}{flag}.",
                mailFolder.FullName, uniqueId, addFlag ? "" : "Not ", Enum.GetName(flag));
        }

        public static async Task<int> DeleteOldMessagesAsync(
            this IMailFolder mailFolder, int days = 28, CancellationToken ct = default)
        {
            if (!mailFolder.IsOpen || mailFolder.Access != FolderAccess.ReadWrite)
                lock (mailFolder.SyncRoot)
                    mailFolder.Open(FolderAccess.ReadWrite, ct);

            var query = SearchQuery.DeliveredBefore(DateTime.Now.Date.AddDays(-days)).And(SearchQuery.Seen);
            IList<UniqueId> matchedUids = await mailFolder.SearchAsync(query, ct);

            //if (Monitor.TryEnter(mailFolder.SyncRoot))
            {
                lock (mailFolder.SyncRoot)
                {
                    mailFolder.AddFlags(matchedUids, MessageFlags.Deleted, true, ct);
                    mailFolder.Close(true); //mailFolder.Expunge(); //, ct
                }
                if (matchedUids.Count > 0)
                    _logger.LogInformation("{0} message{s} more than {2} day{s} old deleted from '{Folder}' ({Count}).",
                        matchedUids.Count, matchedUids.Count.S(), days, days.S(), mailFolder.FullName, mailFolder.Count);
            }

            return matchedUids.Count;
        }
    }
}
