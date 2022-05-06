using MimeKit;
using MimeKit.IO;
using MimeKit.Text;
using MimeKit.Utils;
using Zue.Common;
using MailKitExtensions.Attachments;
using MailKitExtensions.Reader;
using MailKitExtensions.Writer;
using MailKitExtensions.Helpers;

namespace MailKitExtensions.Receiver
{
    public static class ReceiverMimeMessages
    {
        public const string RE = "RE:";
        public const string FW = "FW:";

        private static ILog _logger = LogUtil.GetLogger(nameof(ReceiverMimeMessages));

        public static IEnumerable<MimeMessage> ReplyToSender(this IEnumerable<MimeMessage> originals, string bodyText = "", string from = "", bool replyToAll = false, bool includeEmbedded = true, bool setHtml = true)
        {
            if (originals == null)
                throw new ArgumentNullException(nameof(originals));

            return originals.Select(original => original.BuildReMessage(bodyText, from, null, replyToAll, includeEmbedded, setHtml));
        }

        public static MimeMessage BuildFwMessage(this MimeMessage original, string to, string bodyText = "", string from = "", bool sendToAll = false, bool includeEmbedded = true, bool setHtml = true)
        {
            var mTo = EmailHelper.ParseMailboxAddress(to);
            var mFrom = EmailHelper.ParseMailboxAddress(from);
            var mReplyTo = original.BuildReplyToAddresses(mFrom);
            return original.BuildMimeMessageResponse(mFrom, mTo, FW, bodyText, false, mReplyTo, sendToAll, includeEmbedded, true, setHtml);
        }

        public static MimeMessage BuildReMessage(this MimeMessage original, string bodyText = "", string from = "", string replyTo = "", bool replyToAll = false, bool includeEmbedded = true, bool setHtml = true)
        {
            var mTo = original.BuildReplyAddresses(replyToAll).Mailboxes;
            var mFrom = EmailHelper.ParseMailboxAddress(from);
            var mReplyTo = string.IsNullOrEmpty(replyTo) ? original.BuildReplyToAddresses(mFrom) : EmailHelper.ParseMailboxAddress(replyTo);
            return original.BuildMimeMessageResponse(mFrom, mTo, RE, bodyText, false, mReplyTo, replyToAll, includeEmbedded, true, setHtml);
        }

        public static MimeMessage BuildReturnToSenderMessage(
            this MimeMessage original, string bodyText = "", string descriptionA = "", string descriptionB = "", bool isReply = true, bool toForwarder = false)
        {
            var mTo = original.BuildReplyAddresses(false).Mailboxes;
            var body = bodyText.Replace("{{BodyText}}", isReply ? descriptionA : descriptionB);
            var mReplyTo = original.ResentFrom.Count > 0 ? original.From.Mailboxes : original.To.Mailboxes;
            return original.BuildMimeMessageResponse(null, mTo, RE, body, false, mReplyTo, false, true, true, true);
        }

        public static IEnumerable<MimeMessage> SplitAttachments(
            this MimeMessage original, string to, string subject, string bodyText, IEnumerable<MimeEntity> attachments)
        {
            if (original == null)
                throw new ArgumentNullException(nameof(original));

            var message = new MimeMessage();
            original.BuildIdReferences(ref message);
            message.Body = original.BuildTextPart(bodyText, false, false, false);
            var toAddresses = EmailHelper.ParseMailboxAddress(to);
            var mReplyTo = original.ResentFrom.Count > 0 ? original.From.Mailboxes : original.To.Mailboxes;

            IList<MimeMessage> messages = new List<MimeMessage>();
            if (attachments.IsNotNullOrEmpty())
            {
                foreach (var attachment in attachments)
                {
                    var multipartBody = message.Body.BuildMultipart(attachment);
                    var copy = new MimeMessage(message.From, toAddresses, subject, multipartBody);
                    copy.ReplyTo.AddRange(mReplyTo);
                    copy.InReplyTo = message.InReplyTo;
                    copy.References.AddRange(message.References);
                    messages.Add(copy);
                }
            }
            return messages;
        }

        public static MimeMessage BuildMimeMessageResponse(this MimeMessage original, IEnumerable<MailboxAddress>? from, IEnumerable<MailboxAddress> to,
            string subjectPrefix = "", string bodyText = "", bool replaceSubject = false, IEnumerable<MailboxAddress>? replyTo = null,
            bool sendToAll = false, bool includeEmbedded = true, bool includeAttachments = true, bool setHtml = true)
        {
            if (original == null)
                throw new ArgumentNullException(nameof(original));

            var message = new MimeMessage();

            if (from.IsNotNullOrEmpty())
                message.From.AddRange(from);

            if (replyTo.IsNotNullOrEmpty())
                message.ReplyTo.AddRange(replyTo);

            // include all of the recipients except ourselves
            message.To.AddRange(sendToAll || from == null ? to : to.Excluding(from));

            // set the subject with prefix check
            message.Subject = replaceSubject ? subjectPrefix :
                EmailHelper.GetPrefixedSubject(original.Subject, subjectPrefix);

            // construct the In-Reply-To and References headers
            original.BuildIdReferences(ref message);

            // set the message body
            message.Body = original.BuildTextPart(bodyText, includeEmbedded, includeAttachments, setHtml);

            return message;
        }

        // reply to the sender of the message
        private static InternetAddressList BuildReplyAddresses(this MimeMessage original, bool replyToAll = false)
        {
            if (original == null)
                throw new ArgumentNullException(nameof(original));

            var to = new InternetAddressList();
            if (original.ResentFrom.Count > 0)
            {
                to.AddRange(original.ResentFrom);
            }
            else if (original.ReplyTo.Count > 0)
            {
                to.AddRange(original.ReplyTo);
            }
            else if (original.From.Count > 0)
            {
                to.AddRange(original.From);
            }
            else if (original.Sender != null)
            {
                to.Add(original.Sender);
            }

            if (replyToAll && original.ResentFrom.Count < 1)
            {
                // include all of the other original recipients
                to.AddRange(original.To);
                to.AddRange(original.Cc);
            }

            return to;
        }

        private static IEnumerable<MailboxAddress> BuildReplyToAddresses(
            this MimeMessage original, IEnumerable<MailboxAddress> mFrom)
        {
            return original.ResentFrom.Count > 0 ? original.ResentFrom.Mailboxes :
                original.To.Mailboxes.Contains(mFrom) ? mFrom : original.From.Mailboxes;
        }

        internal static bool Contains(
            this IEnumerable<MailboxAddress> mailboxA, IEnumerable<MailboxAddress> mailboxB)
        {
            return mailboxA.Any(a => !mailboxB.Any(b =>
                a.Address.Equals(b.Address, StringComparison.OrdinalIgnoreCase)));
        }

        internal static IEnumerable<MailboxAddress> Excluding(
            this IEnumerable<MailboxAddress> mailboxA, IEnumerable<MailboxAddress> mailboxB)
        {
            return mailboxA.Where(a => !mailboxB.Any(b =>
                a.Address.Equals(b.Address, StringComparison.OrdinalIgnoreCase)));
        }

        // construct the In-Reply-To and References headers
        private static void BuildIdReferences(this MimeMessage original, ref MimeMessage message)
        {
            if (!string.IsNullOrEmpty(original.MessageId))
            {
                message.InReplyTo = original.MessageId;
                foreach (var id in original.References)
                    message.References.Add(id);
                message.References.Add(original.MessageId);
            }
        }

        // quote the original message text
        internal static MimeEntity BuildTextPart(this MimeMessage original,
            string prependText = "", bool includeEmbedded = true, bool includeAttachments = true, bool setHtml = true)
        {
            if (original == null)
                return new TextPart();

            bool isHtml = setHtml || original.HtmlBody != null;
            var format = isHtml ? TextFormat.Html : TextFormat.Plain;

            var visitor = new ReplyVisitor(includeEmbedded, includeAttachments);
            if (!string.IsNullOrEmpty(prependText))
                visitor.PrependMessage = prependText;
            visitor.Visit(original);

            var mimeBody = visitor.Body != null ? visitor.Body :
                new TextPart(format)
                {
                    Text = original.QuoteOriginal(prependText)
                };

            if (includeAttachments && visitor.Attachments.IsNotNullOrEmpty())
            {
                mimeBody = mimeBody.BuildMultipart(visitor.Attachments.ToArray());
            }

            return mimeBody;
        }

        public static MimeEntity BuildMultipart(
            this MimeEntity mimeBody, params MimeEntity[] attachments)
        {
            if (attachments == null || attachments.Length == 0)
                return mimeBody;

            var multipart = new Multipart("mixed");
            if (mimeBody != null)
                multipart.Add(mimeBody);
            if (attachments.Length > 1)
                multipart.AddRange(attachments);
            else if (attachments.Length == 1)
                multipart.Add(attachments[0]);

            return multipart;
        }

        public static MimeEntity BuildMultipart(
            this MimeEntity mimeBody, IEnumerable<MimeEntity> mimeEntities)
        {
            if (mimeEntities.IsNullOrEmpty())
                return mimeBody;
            var multipart = new Multipart("mixed");
            if (mimeBody != null)
                multipart.Add(mimeBody);
            foreach (var mimeEntity in mimeEntities)
                if (mimeEntity != null)
                    multipart.Add(mimeEntity);
            return multipart;
        }

        private static async Task DownloadAllAttachmentsAsync(
            this IEnumerable<MimeEntity> attachments, string downloadPath, CancellationToken ct = default)
        {
            if (attachments != null && !string.IsNullOrEmpty(downloadPath))
            {
                //_logger.LogDebug("Downloading attachments to '{0}'.", downloadPath);
                foreach (MimePart attachment in attachments)
                {
                    await attachment.DownloadAttachmentAsync(downloadPath, ct);
                }
            }
        }

        private static async Task DownloadAttachmentAsync(
            this MimePart attachment, string downloadPath, CancellationToken ct = default)
        {
            using var stream = FileHandler.GetFileOutputStream(
                downloadPath, attachment.FileName);
            await attachment.WriteToStreamAsync(stream, ct);
            //_logger.LogDebug($"{attachment.FileName} downloaded.");
        }

        public static async Task<MimeMessage> Copy(
            this MimeMessage message, CancellationToken ct = default) =>
            await message.CloneStreamReferences(true, ct);

        public static async Task<MimeMessage> Clone(
            this MimeMessage message, CancellationToken ct = default) =>
            await message.CloneStreamReferences(false, ct);

        private static async Task<MimeMessage> CloneStreamReferences(
            this MimeMessage message, bool persistent, CancellationToken ct = default)
        {
            using var memory = new MemoryBlockStream();
            message.WriteTo(memory);
            memory.Position = 0;
            return await MimeMessage.LoadAsync(memory, persistent, ct);
        }

        internal static string GetOnDateSenderWrote(this MimeMessage message)
        {
            var sender = message.Sender ?? message.From.Mailboxes.FirstOrDefault();
            var name = sender != null ? !string.IsNullOrEmpty(sender.Name) ?
                sender.Name : sender.FormatMailboxName().Name : "someone";

            return string.Format("On {0}, {1} wrote:", message.Date.ToString("f"), name);
        }

        internal static string QuoteText(string text, string prefix = "")
        {
            using (var quoted = new StringWriter())
            {
                if (!string.IsNullOrEmpty(prefix))
                {
                    quoted.WriteLine(prefix);
                }

                if (!string.IsNullOrEmpty(text))
                {
                    using (var reader = new StringReader(text))
                    {
                        string line;

                        while ((line = reader.ReadLine()) != null)
                        {
                            quoted.Write("> ");
                            quoted.WriteLine(line);
                        }
                    }
                }

                return quoted.ToString();
            }
        }

        public static string PrependMessage(this MimeMessage original, string message = "")
        {
            using var text = new StringWriter();
            if (original.HtmlBody == null)
            {
                text.WriteLine(message);
                text.WriteLine();
                text.WriteLine();
                if (!string.IsNullOrEmpty(original.TextBody))
                {
                    text.WriteLine(original.TextBody);
                    text.WriteLine();
                }
            }
            else
            {
                text.WriteLine(message);
                text.WriteLine("<br /><br />");
                if (!string.IsNullOrEmpty(original.HtmlBody))
                {
                    text.WriteLine(original.HtmlBody);
                    text.WriteLine("<br />");
                }
            }

            return text.ToString();
        }

        public static string QuoteOriginal(this MimeMessage original, string message = "")
        {
            //var timestamp = envelope.Date.GetValueOrDefault().ToLocalTime();
            string rfc882DateTime = DateUtils.FormatDate(original.Date);
            using var text = new StringWriter();
            if (original.HtmlBody == null)
            {
                text.WriteLine(message);
                text.WriteLine();
                text.WriteLine("-----Original Message-----");
                text.WriteLine("From: {0}", original.From);
                text.WriteLine("Sent: {0}", rfc882DateTime);
                text.WriteLine("To: {0}", original.To);
                if (!original.Cc.Mailboxes.IsNullOrEmpty())
                    text.WriteLine("Cc: {0}", original.Cc);
                text.WriteLine("Subject: {0}", original.Subject);
                if (!original.Attachments.IsNullOrEmpty())
                    text.WriteLine("Attachment{0}: {1}", original.Attachments.Count().S(),
                        original.Attachments.GetAttachmentNames().ToEnumeratedString());
                text.WriteLine("ID: {0}", original.MessageId);
                if (original.ResentFrom.Mailboxes.IsNotNullOrEmpty())
                    text.WriteLine("Resent From: {0}", original.ResentFrom);
                text.WriteLine();
                if (!string.IsNullOrEmpty(original.TextBody))
                {
                    text.WriteLine(original.TextBody);
                    text.WriteLine();
                }
            }
            else
            {
                //text.WriteLine("<body>");
                text.WriteLine("<div><p>");
                text.WriteLine(message);
                text.WriteLine("</p></div><br />");
                text.WriteLine("<div style=\"border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm\">");
                text.WriteLine("<div><p><hr />");
                //text.WriteLine("<p><hr /><div style='border-left: 1px #ccc solid; margin: 0 0 0 .8ex; padding-left: 1ex;'>");
                //text.WriteLine("<blockquote style='border-left: 1px #ccc solid; margin: 0 0 0 .8ex; padding-left: 1ex;'>");
                text.WriteLine("<b>From:</b> {0}<br />",
                    original.From.FormatMailboxNameAddress());
                text.WriteLine("<b>To:</b> {0}<br />",
                    original.To.FormatMailboxNameAddress());
                if (original.Cc.Mailboxes.IsNotNullOrEmpty())
                    text.WriteLine("<b>Cc:</b> {0}<br />",
                        original.Cc.FormatMailboxNameAddress());
                text.WriteLine("<b>Subject:</b> {0}<br />", original.Subject);
                if (original.Attachments.IsNotNullOrEmpty())
                    text.WriteLine("<b>Attachment{0}:</b> {1}<br />", original.Attachments.Count().S(),
                        original.Attachments.GetAttachmentNames().ToEnumeratedString());
                text.WriteLine("<b>ID:</b> {0}<br />", original.MessageId);
                if (original.ResentFrom.Mailboxes.IsNotNullOrEmpty())
                    text.WriteLine("<b>Resent From:</b> {0}<br />",
                        original.ResentFrom.FormatMailboxNameAddress());
                text.WriteLine("</p></div><br />");
                if (!string.IsNullOrEmpty(original.HtmlBody))
                {
                    text.WriteLine("<div>");
                    text.WriteLine(original.HtmlBody);
                    text.WriteLine("</div><br />");
                }
                //text.WriteLine("</blockquote>");
                //text.WriteLine("</body>");
            }

            return text.ToString();
        }

        public static string DecodeHtmlBody(string html)
        {
            if (html == null)
                return null;

            bool previousWasNewLine = false;
            using var writer = new StringWriter();
            using (var reader = new StringReader(html))
            {
                var tokenizer = new HtmlTokenizer(reader)
                {
                    DecodeCharacterReferences = true
                };

                while (tokenizer.ReadNextToken(out HtmlToken token))
                {
                    switch (token.Kind)
                    {
                        case HtmlTokenKind.Data:
                            var data = token as HtmlDataToken;
                            if (!string.IsNullOrWhiteSpace(data?.Data) &&
                                !data.Data.StartsWith("<!--"))
                            {
                                writer.Write(data.Data);
                                if (!data.Data.EndsWith(Environment.NewLine))
                                    previousWasNewLine = false;
                            }
                            break;
                        case HtmlTokenKind.Tag:
                            var tag = (HtmlTagToken)token;
                            switch (tag.Id)
                            {
                                case HtmlTagId.BlockQuote:
                                case HtmlTagId.Br:
                                    if (!previousWasNewLine)
                                    {
                                        writer.Write(Environment.NewLine);
                                        previousWasNewLine = true;
                                    }
                                    break;
                                case HtmlTagId.P:
                                    if (!previousWasNewLine &&
                                        tag.IsEndTag || tag.IsEmptyElement)
                                    {
                                        writer.Write(Environment.NewLine);
                                        previousWasNewLine = true;
                                    }
                                    break;
                            }
                            break;
                    }
                }
            }

            return writer.ToString();
        }

        public static IList<string> DecodeHtmlHrefs(string html)
        {
            if (html == null)
                return null;

            var hrefs = new List<string>();
            using (var reader = new StringReader(html))
            {
                var tokenizer = new HtmlTokenizer(reader)
                {
                    DecodeCharacterReferences = true
                };

                while (tokenizer.ReadNextToken(out HtmlToken token))
                {
                    if (token.Kind == HtmlTokenKind.Tag &&
                        token is HtmlTagToken tag &&
                        tag.Id == HtmlTagId.A &&
                        tag.Attributes != null &&
                        tag.Attributes.Any())
                    {
                        hrefs.AddRange(tag.Attributes.Select(a => a.Value));
                    }
                }
            }

            return hrefs;
        }

        public static IEnumerable<Uri> GetUris(string bodyText, bool isHtml)
        {
            IEnumerable<string>? results = null;
            try
            {
                results = isHtml ? DecodeHtmlHrefs(bodyText) : GetUrls(bodyText);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to parse URLs from the email body.");
            }
            return results.ConvertToUris().RemoveNulls();
        }

        private static IEnumerable<string>? GetUrls(string bodyText)
        {
            char[] splitChars = new char[] { ' ', '\r', '\n', ',', '\u00A0' }; // '\u00A0' = (char)160 = &nbsp
            return bodyText?.Split(splitChars, StringSplitOptions.RemoveEmptyEntries).FilterHrefs();
        }

        private static IEnumerable<string> FilterHrefs(this IEnumerable<string> htmlHrefs)
        {
            //if (htmlHrefs.Any(a => a.Contains("pay/#/"))) System.Diagnostics.Debugger.Break();
            return htmlHrefs?.Where(a => !string.IsNullOrEmpty(a) && a.Contains("https://", StringComparison.OrdinalIgnoreCase))
                .Select(href => href.Replace("#/", "%23/")) ?? Array.Empty<string>(); //MYOB doesn't follow RFC 1738 '\u0023'
        }

        private static IEnumerable<Uri?> ConvertToUris(this IEnumerable<string> htmlHrefs)
        {
            return htmlHrefs.FilterHrefs().Select(a => Uri.TryCreate(a, UriKind.Absolute, out Uri? result) ? result : null);
        }

        private static IEnumerable<Uri> RemoveNulls(this IEnumerable<Uri?>? uris)
        {
            return uris?.Where(a => a?.Segments?.Length > 1) ?? Array.Empty<Uri>();
        }

        public static void Save(this MimeMessage message, uint id, bool createDirectory = false, params string[] downloadPath)
        {
            if (message.From.Count == 0)
                message.From.Add(MailKitExtensions.Sender.EmailSender.DefaultFromAddress);

            string path = downloadPath?.Length > 1 ? Path.Combine(downloadPath) :
                downloadPath?.Length > 0 ? downloadPath[0] : "";
            if (createDirectory)
                Directory.CreateDirectory(path);
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
            {
                var fileNames = FileHandler.EnumerateFilesFromFolder(path, "*.eml");
                if (!fileNames.Any(f => f.StartsWith(id.ToString())))
                    message.Save(Path.Combine(path, $"{id}_{message.MessageId}.eml"));
            }
        }

        public static void Save(this MimeMessage message, string name = "message.eml", bool dosFormat = false)
        {
            if (dosFormat)
            {
                // clone the default formatting options
                var format = FormatOptions.Default.Clone();

                // override the line-endings to be DOS no matter what platform we are on
                format.NewLineFormat = NewLineFormat.Dos;

                message.WriteTo(format, name);
            }
            else
            {
                message.WriteTo(name);
            }

//#if DEBUG
//            try
//            {
//                string folderPath = Path.Combine(@"C:\Temp\Emails\", "Text");
//                Directory.CreateDirectory(folderPath);
//                string filePath = Path.Combine(folderPath, message.MessageId + ".txt");
//                using (var writer = new StreamWriter(filePath, true))
//                {
//                    writer.Write(DecodeHtmlBody(message.HtmlBody));
//                }
//            }
//            catch (IOException ex)
//            {
//                System.Diagnostics.Debugger.Break();
//            }
//            //await File.WriteAllLinesAsync("WriteLines.txt", lines);
//            //region = original.ResentFrom; TODO use this
//#endif
        }
    }
}
