using MailKit;
using MimeKit;
using MimeKit.Text;
using System.Diagnostics;
using MailKitExtensions.Attachments;
using Zue.Common;
using MailKitExtensions.Helpers;

namespace MailKitExtensions.Writer
{
    public static class EmailHelper
    {
        private static ILog _logger = LogUtil.GetLogger(nameof(EmailHelper));

        public static MimeMessage CreateMimeMessage(
            string from, string to, string subject, string bodyText,
            bool isHtml = true, string replyTo = "", params string[] attachmentNames)
        {
            var attachments = attachmentNames.GetMimeEntitiesFromFilePaths();
            return CreateMimeMessage(from, to, subject, bodyText, isHtml, replyTo, attachments);
        }

        public static MimeMessage CreateMimeMessage(
            string from, string to, string subject, string bodyText,
            bool isHtml = true, string replyTo = "", IEnumerable<MimeEntity> attachments = null)
        {
            CreateMimeEnvelope(from, to, bodyText, isHtml, replyTo,
                out IEnumerable<MailboxAddress> mFrom, out IEnumerable<MailboxAddress> mTo,
                out MimeEntity mBody, out IEnumerable<MailboxAddress> mReplyTo);
            return CreateMimeMessage(mFrom, mTo, subject, mBody, mReplyTo, attachments);
        }

        internal static void CreateMimeEnvelope(
            string from, string to, string bodyText, bool isHtml, string replyTo,
            out IEnumerable<MailboxAddress> mFrom, out IEnumerable<MailboxAddress> mTo,
            out MimeEntity mTextBody, out IEnumerable<MailboxAddress> mReplyTo)
        {
            mFrom = ParseMailboxAddress(from);
            mTo = ParseMailboxAddress(to);
            mReplyTo = ParseMailboxAddress(replyTo);
            var format = isHtml ? TextFormat.Html : TextFormat.Plain;
            mTextBody = new TextPart(format) { Text = bodyText ?? "" };
        }

        public static IEnumerable<MailboxAddress> ParseMailboxAddress(string value)
        {
            char[] replace = new char[] { '_', '.', '-' };
            char[] separator = new char[] { ';', ',', ' ', '|' };
            return string.IsNullOrEmpty(value) ? Array.Empty<MailboxAddress>() :
                value.Split(separator, StringSplitOptions.RemoveEmptyEntries)
                .Select(f => new MailboxAddress(f.Split('@').FirstOrDefault()
                    ?.ToSpaceReplaceTitleCase(replace) ?? f, f));
        }

        internal static string GetPrefixedSubject(string originalSubject, string prefix = "")
        {
            string subject = originalSubject ?? "";
            if (!string.IsNullOrEmpty(prefix) &&
                !subject.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                subject = string.Join(' ', prefix, subject);
            return subject ?? "";
        }

        internal static MimeMessage CreateMimeMessage(
            IEnumerable<MailboxAddress> from, IEnumerable<MailboxAddress> to, string subject,
            MimeEntity textBody, IEnumerable<MailboxAddress> replyTo = null, IEnumerable<MimeEntity> attachments = null)
        {
            var body = textBody ?? new TextPart(TextFormat.Html);
            string attachmentNames = string.Empty;

            if (attachments != null && attachments.Any())
            {
                var multipart = new Multipart("mixed");
                multipart.Add(textBody);
                foreach (var attachment in attachments)
                    multipart.Add(attachment);
                body = multipart;
                attachmentNames = string.Format(", with attached: '{0}'",
                    attachments.GetAttachmentNames().ToEnumeratedString("', '"));
            }

            if (replyTo.IsNullOrEmpty())
            {
                replyTo = from;
            }
            from = from.FormatMailboxName();
            to = to.FormatMailboxName();
            replyTo = replyTo.FormatMailboxName();

            var message = new MimeMessage(from, to, subject, body);
            message.ReplyTo.AddRange(replyTo);
            
            if (_logger != null)
                _logger.LogDebug("Sending to: {0}, subject: '{1}'{2}", to, subject, attachmentNames);
            else
                Trace.TraceInformation("Sending to: {0}, subject: '{1}'{2}", to, subject, attachmentNames);

            return message;
        }

        public static IEnumerable<string> GetEmlFilesFromFolder(params string[] folderPath)
        {
            string path = folderPath?.Length > 1 ? Path.Combine(folderPath) :
                folderPath?.Length > 0 ? folderPath[0] : "";
            IEnumerable<string> uncFileNames = Array.Empty<string>();
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                uncFileNames = FileHandler.EnumerateFilesFromFolder(path, "*.eml", false);
                if (!uncFileNames.Any())
                    uncFileNames = FileHandler.EnumerateFilesFromFolder(path, "*.eml", true);
            }
            return uncFileNames;
        }

        public static IEnumerable<MimeMessage> GetMimeMessagesFromFolder(
            string folderPath, CancellationToken ct = default)
        {
            var emlFiles = GetEmlFilesFromFolder(folderPath);
            return GetMimeMessagesFromPath(emlFiles, ct);
        }

        public static IEnumerable<MimeMessage> GetMimeMessagesFromPath(
            IEnumerable<string> uncFileNames, CancellationToken ct = default)
        {
            if (uncFileNames != null)
            {
                foreach (var filePath in uncFileNames)
                {
                    if (ct.IsCancellationRequested)
                        break;
                    if (File.Exists(filePath))
                    {
                        yield return MimeMessage.Load(filePath, ct);
                    }
                }
            }
        }
    }
}
