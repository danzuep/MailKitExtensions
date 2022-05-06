using MimeKit;
using MailKit;
using MailKit.Search;
using MailKit.Net.Imap;
using MailKit.Security;
using System.Net;
using Microsoft.Extensions.Logging;
using Zue.Common;

namespace MailKitExtensions.Receiver
{
    public static class EmailReceiverExtensions
    {
        private static ILogger _logger = EmailReceiver._logger;

        #region ImapClient
        public static void Connect(this ImapClient client, string host,
            NetworkCredential credential, int port = 993, bool useSsl = true, CancellationToken ct = default)
        {
            if (client is null)
                throw new ArgumentNullException(nameof(client));
            if (string.IsNullOrEmpty(host))
                throw new ArgumentNullException(nameof(host));
            if (credential is null)
                throw new ArgumentNullException(nameof(credential));

            //bool useSsl = credential == null ? false : true;
            //// IMAP port 143 unencrypted, 993 encrypted
            //int port = useSsl ? 993 : 143;

            try
            {
                if (!client.IsConnected && !string.IsNullOrEmpty(host))
                {
                    lock (client.SyncRoot)
                    {
                        client.Connect(host, port, useSsl, ct);
                        if (client.Capabilities.HasFlag(ImapCapabilities.Compress))
                            client.Compress(ct);
                    }
                }
                if (!client.IsAuthenticated && credential != null)
                {
                    if (client.AuthenticationMechanisms.Contains("NTLM"))
                    {
                        var ntlm = new SaslMechanismNtlm(credential);
                        if (ntlm?.Workstation is not null)
                            lock (client.SyncRoot)
                                client.Authenticate(ntlm);
                        else
                            lock (client.SyncRoot)
                                client.Authenticate(credential, ct);
                    }
                    else
                    {
                        lock (client.SyncRoot)
                            client.Authenticate(credential, ct);
                    }
                }
//#if DEBUG
//                _ = client.GetFolderList();
//#endif
            }
            catch (AuthenticationException ex)
            {
                _logger.LogError(ex, "Failed to authenticate.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to connect to email client.");
            }
        }
        #endregion

        internal static IList<string> GetFolderList(this ImapClient client)
        {
            IList<string> mailFolderNames = new List<string>();
            if (client != null && client.IsAuthenticated)
            {
                if (client.PersonalNamespaces.Count > 0)
                {
                    lock (client.SyncRoot)
                    {
                        var rootFolder = client.GetFolder(client.PersonalNamespaces[0]);
                        var subfolders = rootFolder.GetSubfolders().Select(f => f.Name);
                        var inboxSubfolders = client.Inbox.GetSubfolders().Select(f => f.FullName);
                        mailFolderNames.AddRange(inboxSubfolders);
                        mailFolderNames.AddRange(subfolders);
                        _logger.LogDebug("{0} Inbox folders: {1}", subfolders.Count(),
                            inboxSubfolders.ToEnumeratedString());
                        _logger.LogDebug("{0} personal folders: {1}", subfolders.Count(),
                            subfolders.ToEnumeratedString());
                    }
                }
                if (client.SharedNamespaces.Count > 0)
                {
                    lock (client.SyncRoot)
                    {
                        var rootFolder = client.GetFolder(client.SharedNamespaces[0]);
                        var subfolders = rootFolder.GetSubfolders().Select(f => f.Name);
                        mailFolderNames.AddRange(subfolders);
                        _logger.LogDebug("{0} shared folders: {1}", subfolders.Count(),
                            subfolders.ToEnumeratedString());
                    }
                }
                if (client.OtherNamespaces.Count > 0)
                {
                    lock (client.SyncRoot)
                    {
                        var rootFolder = client.GetFolder(client.OtherNamespaces[0]);
                        var subfolders = rootFolder.GetSubfolders().Select(f => f.Name);
                        mailFolderNames.AddRange(subfolders);
                        _logger.LogDebug("{0} other folders: {1}", subfolders.Count(),
                            subfolders.ToEnumeratedString());
                    }
                }
            }
            return mailFolderNames;
        }

        public static IMailFolder ConnectFolder(this IMailFolder mailFolder, bool enableWrite = false, CancellationToken ct = default)
        {
            if (mailFolder == null)
                throw new ArgumentNullException(nameof(mailFolder));

            if (!mailFolder.IsOpen)
            {
                var folderAccess = enableWrite ? FolderAccess.ReadWrite : FolderAccess.ReadOnly;
                lock (mailFolder.SyncRoot)
                    mailFolder.Open(folderAccess, ct);
                _logger.LogTrace("'{0}' ({1}) mail folder opened with {2} access.",
                    mailFolder.FullName, mailFolder.Count, Enum.GetName(folderAccess));
            }
            else if (enableWrite && mailFolder.Access != FolderAccess.ReadWrite)
            {
                _logger.LogTrace("'{0}' ({1}) mail folder SyncRoot changed for ReadWrite access.",
                    mailFolder.FullName, mailFolder.Count);
                lock (mailFolder.SyncRoot)
                    mailFolder.Open(FolderAccess.ReadWrite, ct);
            }

            return mailFolder;
        }

        public static IMailFolder OpenFolder(
            this ImapClient client, string folderName = "INBOX", bool openFolder = true, bool enableWrite = false, CancellationToken ct = default)
        {
            if (client == null)
                throw new ArgumentNullException(nameof(client));

            IMailFolder? mailFolder = null;
            try
            {
                if (string.IsNullOrEmpty(folderName))
                    folderName = "INBOX";
                lock (client.SyncRoot)
                    mailFolder = client.GetFolder(folderName, ct);
            }
            catch (FolderNotFoundException ex)
            {
                _logger.LogWarning(ex, "Mail folder '{0}' not found, defaulting to INBOX.", folderName);
            }
            
            if (mailFolder == null)
                lock (client.SyncRoot)
                    mailFolder = client.Inbox;

            if (openFolder)
                mailFolder.ConnectFolder(enableWrite, ct);

            return mailFolder;
        }

        public static async Task<IList<IMessageSummary>> GetMessageIdsAsync(
            this EmailReceiver emailReceiver, CancellationToken ct = default)
        {
            return await emailReceiver.GetMessageSummariesAsync(
                MessageSummaryItems.UniqueId, ct);
        }

        public static async Task<IList<IMessageSummary>> GetMessageDatesAsync(
            this EmailReceiver emailReceiver, CancellationToken ct = default)
        {
            return await emailReceiver.GetMessageSummariesAsync(
                MessageSummaryItems.UniqueId | MessageSummaryItems.InternalDate, ct);
        }

        public static async Task<IList<IMessageSummary>> GetMessageSummariesAsync(
            this EmailReceiver emailReceiver, MessageSummaryItems filter = MessageSummaryItems.UniqueId, CancellationToken ct = default)
        {
            if (emailReceiver == null)
                throw new ArgumentNullException(nameof(emailReceiver));
            var mailFolder = emailReceiver.ConnectFolder(false, ct);
            return await mailFolder.FetchAsync(0, -1, filter, ct);
        }

        public static async Task<IList<IMessageSummary>> GetMessageSummariesAsync(
            this EmailReceiver emailReceiver, IEnumerable<UniqueId> uniqueIds, MessageSummaryItems filter = MessageSummaryItems.UniqueId, CancellationToken ct = default)
        {
            if (emailReceiver == null)
                throw new ArgumentNullException(nameof(emailReceiver));
            if (uniqueIds == null)
                throw new ArgumentNullException(nameof(uniqueIds));
            var mailFolder = emailReceiver.ConnectFolder(false, ct);
            return await mailFolder.FetchAsync(uniqueIds.ToList(), filter, ct);
        }

        public static async Task<IList<IMessageSummary>> GetMessageSummariesAsync(
            this IMailFolder mailFolder, uint from, uint to = 0, MessageSummaryItems filter = MessageSummaryItems.UniqueId, CancellationToken ct = default)
        {
            if (to < from)
                to = from;
            if (mailFolder == null)
                throw new ArgumentNullException(nameof(mailFolder));
            bool closeWhenFinished = !mailFolder.IsOpen;
            if (closeWhenFinished)
                await mailFolder.OpenAsync(FolderAccess.ReadOnly, ct);
            var messageSummaries = await mailFolder.FetchAsync(0, -1, filter, ct);
            var filteredSummaries = messageSummaries.Where(msg => msg.UniqueId.Id >= from && msg.UniqueId.Id <= to).ToList();
            if (closeWhenFinished)
                await mailFolder.CloseAsync(false, ct);
            return filteredSummaries;
        }

        public static async Task<IEnumerable<IMessageSummary>> GetEmailRangeAsync(
            this EmailReceiver emailReceiver, DateTimeOffset from, DateTimeOffset to, CancellationToken ct = default)
        {
            //Note: IMessageSummary time may be 60 seconds after the MimeMessage time
            //_logger.LogDebug("Searching for message summaries where date > '{0}' and date < '{1}'...", from.Date, to.Date);
            var messageSummaries = await emailReceiver.GetMessageDatesAsync(ct);
            var filteredSummaries = messageSummaries.Where(msg => msg.Date >= from && msg.Date <= to);
            //_logger.LogDebug("'{0}' ({1}) returned {2} matches between '{3}' and '{4}'.",
            //    emailReceiver.MailFolder.FullName, messageSummaries.Count, filteredSummaries.Count(), from, to);
            //_logger.LogDebug("{0} message summaries found between '{1}' and '{2}' (inclusive).", filteredSummaries.Count(), from, to);
            return filteredSummaries;
        }

        public static async Task<IList<IMessageSummary>> SearchMessageSummariesAsync(
            this EmailReceiver emailReceiver, string keywords, bool includeBody = true,
            DateTime? deliveredAfter = null, DateTime? deliveredBefore = null, CancellationToken ct = default)
        {
            if (emailReceiver == null)
                throw new ArgumentNullException(nameof(emailReceiver));
            var mailFolder = emailReceiver.ConnectFolder(false, ct);
            SearchQuery query = SearchQuery.All;
            if (!string.IsNullOrWhiteSpace(keywords))
            {
                var subjectQuery = BuildSubjectSearchQuery(keywords);
                var bodyQuery = BuildBodySearchQuery(keywords);
                query = includeBody ? subjectQuery.Or(bodyQuery) : subjectQuery;
            }
            if (deliveredAfter != null)
            {
                var dateQuery = BuildDateSearchQuery(deliveredAfter.Value, deliveredBefore);
                query = dateQuery.And(query);
            }
            var matchedUids = await mailFolder.SearchAsync(query, ct);

            return await mailFolder.FetchAsync(matchedUids, EmailReceiver.ItemFilter, ct);
        }

        public static async Task<IList<IMessageSummary>> SearchMessageSummariesAsync(
            this EmailReceiver emailReceiver, bool includeBody, string keywords,
            DateTimeOffset? deliveredAfter = null, DateTimeOffset? deliveredBefore = null, CancellationToken ct = default)
        {
            DateTime? after = deliveredAfter?.DateTime.AddDays(-1);
            DateTime? before = deliveredBefore?.DateTime.AddDays(1);
            string keyword = keywords ?? "";
            bool isSplit = keyword.Contains('|');
            bool hasPeriod = keyword.Contains('.');
            var splitKeywords = keyword.Split('|', StringSplitOptions.RemoveEmptyEntries);
            if (isSplit && hasPeriod)
            {
                var splitUrls = splitKeywords.Select(k => k.Split('.', StringSplitOptions.RemoveEmptyEntries));
                var searchTerms = splitUrls.Select(m => string.Join('.', m.Length > 2 ? m.Skip(m.Length - 2) : m));
                keyword = string.Join('|', searchTerms);
            }
            _logger.LogDebug("Searching message summaries for '{0}'{3}, date > '{1:yyyy-MM-dd}' and date < '{2:yyyy-MM-dd}'...",
                keyword, isSplit ? " (keywords split on '|')" : "", after?.Date, before?.Date);
            var messageSummaries = await emailReceiver.SearchMessageSummariesAsync(keyword, includeBody, after, before, ct);
            int messageCount = messageSummaries.Count;
            if (deliveredAfter != null && messageSummaries.Count > 0)
                messageSummaries = messageSummaries.Where(msg => msg.InternalDate >= deliveredAfter &&
                    msg.InternalDate <= (deliveredBefore ?? DateTimeOffset.Now)).ToList();
            if (isSplit && hasPeriod && messageSummaries.Count > 0)
                messageSummaries = messageSummaries.Where(msg => splitKeywords.Any(a =>
                    msg?.GetMailBodyTextAsync(ct).GetAwaiter().GetResult().Contains(a, StringComparison.OrdinalIgnoreCase) ?? false)).ToList();
            _logger.LogDebug("'{0}' returned {1} matches for '{2}' between '{3}' and '{4}' (down from {5}).",
                emailReceiver.MailFolder.FullName, messageSummaries.Count, keywords, deliveredAfter, deliveredBefore, messageCount);
            return messageSummaries;
        }

        internal static SearchQuery BuildSubjectSearchQuery(string keywords)
        {
            // Split string into a list of queries and enumerate
            return keywords?.Split('|')
                ?.Select(key => SearchQuery.SubjectContains(key))
                ?.ToList()?.EnumerateOr() ??
                // return 'false' if null
                SearchQuery.Recent.And(SearchQuery.Old);
        }

        internal static SearchQuery BuildBodySearchQuery(string keywords)
        {
            // Split string into a list of queries and enumerate
            return keywords?.Split('|')
                ?.Select(key => SearchQuery.BodyContains(key))
                ?.ToList()?.EnumerateOr() ??
                // return 'false' if null
                SearchQuery.Recent.And(SearchQuery.Old);
        }

        internal static SearchQuery BuildDateSearchQuery(DateTime deliveredAfter, DateTime? deliveredBefore = null)
        {
            DateTime before = deliveredBefore != null ? deliveredBefore.Value : DateTime.Now;
            return SearchQuery.DeliveredAfter(deliveredAfter).And(SearchQuery.DeliveredBefore(before));
        }

        private static SearchQuery EnumerateOr<T>(
            this IList<T> queries) where T : SearchQuery
        {
            T query = queries.FirstOrDefault();

            if (queries?.Count > 1)
            {
                queries.Remove(query);
                // recursively return an 'Or' query
                return query.Or(EnumerateOr(queries));
            }

            return query;
        }

        internal static IEnumerable<Uri> GetUrisFromBody(this IMessageSummary messageSummary, CancellationToken ct = default) =>
            ReceiverMimeMessages.GetUris(messageSummary.GetMailBodyTextAsync(ct).GetAwaiter().GetResult(), messageSummary.HtmlBody != null);

        public static async Task<IEnumerable<Uri>> GetBodyUris(
            this EmailReceiver emailReceiver, UniqueId? messageId, CancellationToken ct = default)
        {
            if (emailReceiver == null)
                throw new ArgumentNullException(nameof(emailReceiver));
            if (messageId == null)
                throw new ArgumentNullException(nameof(messageId));
            var mailFolder = emailReceiver.ConnectFolder(false, ct);
            var messageBodySummaries = await mailFolder.FetchAsync(
                new UniqueId[] { messageId.Value }, MessageSummaryItems.Body, ct);
            return messageBodySummaries.SelectMany(body => body.GetUrisFromBody(ct));
        }

        public static void SaveMimeMessages(
            this IMailFolder mailFolder, IEnumerable<UniqueId> uniqueIds, string downloadPath, CancellationToken ct = default)
        {
            downloadPath = Path.Combine(downloadPath, mailFolder.FullName.Replace("/", "-"));
            bool isEmpty = uniqueIds == null || !uniqueIds.Any();
            if (!isEmpty && !Directory.Exists(downloadPath))
            {
#if DEBUG
                Directory.CreateDirectory(downloadPath);
#else
                _logger.LogWarning("'{0}' does not exist.", downloadPath);
#endif
            }

            if (isEmpty)
                _logger.LogInformation("No Mime UniqueIds were provided.");
            else
            {
                _logger.LogDebug("Saving {0} to '{1}'...", uniqueIds.Count(), downloadPath);
                foreach (var id in uniqueIds)
                {
                    var msg = mailFolder.GetMessage(id, ct);
                    string fileName = $"{id.Id}_{msg.MessageId}.eml";
                    string filePath = Path.Combine(downloadPath, fileName);
                    msg.Save(filePath);
                }
            }
        }

        public static void SaveMimeMessages(
            this EmailReceiver emailReceiver, IEnumerable<UniqueId> uniqueIds, string downloadPath, CancellationToken ct = default)
        {
            emailReceiver.Reconnect();
            SaveMimeMessages(emailReceiver.MailFolder, uniqueIds, downloadPath, ct);
        }

        //public static void SaveMimeMessages(
        //    this EmailReceiver emailReceiver, IEnumerable<MimeMessage> messages, string downloadPath, bool createDirectory = false)
        //{
        //    emailReceiver.Reconnect();
        //    if (createDirectory) // && !Directory.Exists(downloadPath)
        //        Directory.CreateDirectory(downloadPath);
        
        //    Func<uint, string, string> GetEmlPath = (id, description) =>
        //        Path.Combine(downloadPath, $"{id}_{description}.eml");
        //    foreach (var msg in messages)
        //    {
        //        string fileName = $"{msg.Date.Ticks}_{msg.MessageId}.eml";
        //        string filePath = Path.Combine(downloadPath, fileName);
        //        msg.Save(filePath);
        //    }
        //}

        public static void SaveMimeMessages(
            this IEnumerable<MimeMessage> messages, string downloadPath, bool createDirectory = false)
        {
            if (createDirectory) // && !Directory.Exists(downloadPath)
                Directory.CreateDirectory(downloadPath);

            if (messages != null)
            {
                foreach (var msg in messages)
                {
                    if (msg != null)
                    {
                        string fileName = $"{msg.Date.Ticks}_{msg.MessageId}.eml";
                        string filePath = Path.Combine(downloadPath, fileName);
                        msg.Save(filePath);
                    }
                }
            }
        }

        public static async ValueTask<MimeMessage> GetMimeMessageAsync(this EmailReceiver emailReceiver,
            UniqueId uniqueId, string folderName = "INBOX", bool openFolder = true, bool enableWrite = false, CancellationToken ct = default)
        {
            var mailFolder = emailReceiver.OpenFolder(folderName, openFolder, enableWrite, ct);
            return await uniqueId.GetMimeMessageAsync(mailFolder, ct);
        }

        public static async ValueTask<MimeMessage> GetMimeMessageAsync(
            this UniqueId uniqueId, IMailFolder mailFolder, CancellationToken ct = default)
        {
            MimeMessage? mimeMessage = null;

            if (mailFolder != null && uniqueId.IsValid)
            {
                try
                {
                    mailFolder.ConnectFolder(false, ct);
                    mimeMessage = await mailFolder.GetMessageAsync(uniqueId, ct);
                }
                catch (MessageNotFoundException ex)
                {
                    _logger.LogWarning(ex, "'{0}' {1} not found.", mailFolder.FullName, uniqueId);
                }
                catch (ImapCommandException ex)
                {
                    _logger.LogWarning(ex, "'{0}' {1} was moved before it could be downloaded.", mailFolder.FullName, uniqueId);
                }
                catch (ImapProtocolException ex)
                {
                    _logger.LogWarning(ex, "'{0}' {1} not downloaded.", mailFolder.FullName, uniqueId);
                }
                catch (IOException ex)
                {
                    _logger.LogWarning(ex, "'{0}' {1} not downloaded.", mailFolder.FullName, uniqueId);
                }
                catch (InvalidOperationException ex) // includes FolderNotOpenException
                {
                    _logger.LogWarning(ex, "'{0}' {1} not downloaded.", mailFolder.FullName, uniqueId);
                }
                catch (OperationCanceledException) // includes TaskCanceledException
                {
                    _logger.LogDebug("Message download task was cancelled.");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "'{0}' {1} failed to return a valid MimeMessage.", mailFolder.FullName, uniqueId);
                }
            }

            return mimeMessage ?? new MimeMessage();
        }

        public static MimeMessage GetMimeMessage(
            this IMessageSummary mail, CancellationToken ct = default)
        {
            MimeMessage? mimeMessage = null;

            if (mail?.Folder != null)
            {
                try
                {
                    mail.Folder.ConnectFolder(false, ct);
                    lock (mail.Folder.SyncRoot)
                        mimeMessage = mail.Folder.GetMessage(mail.UniqueId, ct);
                }
                catch (MessageNotFoundException ex)
                {
                    _logger.LogWarning("{0}: {1} was moved before it could be downloaded. {2}", mail.Folder.FullName, mail.UniqueId, ex.Message);
                }
                catch (ImapCommandException ex)
                {
                    _logger.LogWarning("{0}: {1} was moved before it could be downloaded. {2}", mail.Folder.FullName, mail.UniqueId, ex.Message);
                }
                catch (ImapProtocolException ex)
                {
                    _logger.LogWarning("{0}: {1} not downloaded. {2}", mail.Folder.FullName, mail.UniqueId, ex.Message);
                }
                catch (IOException ex)
                {
                    _logger.LogWarning("{0}: {1} not downloaded. {2}", mail.Folder.FullName, mail.UniqueId, ex.Message);
                }
                catch (InvalidOperationException ex) // includes FolderNotOpenException
                {
                    _logger.LogWarning("{0}: {1} not downloaded. {2}", mail.Folder.FullName, mail.UniqueId, ex.Message);
                }
                catch (OperationCanceledException) // includes TaskCanceledException
                {
                    _logger.LogDebug("Message download task was cancelled.");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed get MimeMessage from IMessageSummary.");
                }
            }

            return mimeMessage ?? new MimeMessage();
        }

        public static IEnumerable<MimeMessage> GetMimeMessages(
            this EmailReceiver emailReceiver, IEnumerable<UniqueId> uniqueIds, CancellationToken ct = default)
        {
            emailReceiver.Reconnect();
            return emailReceiver.MailFolder.GetMimeMessages(uniqueIds, ct);
        }

        public static IEnumerable<MimeMessage> GetMimeMessages(
            this IMailFolder mailFolder, IEnumerable<UniqueId> uniqueIds, CancellationToken ct = default)
        {
            if (mailFolder != null && uniqueIds != null && uniqueIds.Any())
            {
                mailFolder.ConnectFolder(false, ct);
                foreach (var id in uniqueIds)
                {
                    var message = mailFolder.GetMessage(id, ct);

                    if (message != null)
                        yield return message;

                    if (mailFolder.Access == FolderAccess.ReadWrite)
                        mailFolder.FlagMessage(id, MessageFlags.Seen, true, ct);
                }
            }
        }

        public static async Task<IList<MimeMessage>> GetMimeMessagesAsync(
            this IMailFolder mailFolder, int limit = 10, int offset = 0, CancellationToken ct = default)
        {
            IList<MimeMessage> mimeMessages = new List<MimeMessage>();
            if (mailFolder != null)
            {
                if (!mailFolder.IsOpen)
                    await mailFolder.OpenAsync(FolderAccess.ReadOnly, ct);

                int start = (mailFolder.Count < limit + offset) ? offset : 0;
                int count = (mailFolder.Count > limit) ? limit : mailFolder.Count;
                for (int i = start; i < count; i++)
                {
                    var message = await mailFolder.GetMessageAsync(i, ct);

                    if (message != null)
                        mimeMessages.Add(message);

                    if (mailFolder.Access == FolderAccess.ReadWrite)
                        await mailFolder.AddFlagsAsync(i, MessageFlags.Seen, true, ct);
                }
            }
            return mimeMessages;
        }

        public static async Task<int> DeleteExpiredAsync(
            this EmailReceiver emailReceiver, int days, IList<string> folderNames, CancellationToken ct = default)
        {
            int deleteCount = 0;
            if (emailReceiver is not null && !ct.IsCancellationRequested && days > 0 &&
                folderNames?.Count > 0 && !string.IsNullOrEmpty(folderNames[0]))
            {
                emailReceiver.Reconnect(ct);
                string originalMailFolder = emailReceiver.MailFolder?.FullName ?? "INBOX";
                foreach (var folder in folderNames)
                {
                    if (ct.IsCancellationRequested)
                        break;
                    var mailFolder = emailReceiver.OpenFolder(folder, true, true);
                    deleteCount = await mailFolder.DeleteOldMessagesAsync(days);
                }
                string currentMailFolder = emailReceiver.MailFolder?.FullName ?? "INBOX";
                bool isDifferentFolder = currentMailFolder != originalMailFolder;
                if (isDifferentFolder)
                {
                    emailReceiver.OpenFolder(originalMailFolder, false, false);
                    _logger.LogTrace("MailFolder changed back to '{0}'.", originalMailFolder);
                }
            }
            return deleteCount;
        }
    }
}
