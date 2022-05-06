using MailKit;
using MimeKit;
using MimeKit.Text;
using System.Diagnostics;
using Microsoft.Extensions.Logging;
using MailKitExtensions.Receiver;
using MailKitExtensions.Reader;
using MailKitExtensions.Writer;
using MailKitExtensions.Helpers;
using Zue.Common;

namespace MailKitExtensions
{
    public static class MimeMessageExtensions
    {
        public const string RE = "RE:";
        public const string FW = "FW:";

        private static ILogger _logger = EmailReceiver._logger;

        public static MimeMessage ModifyMessage(
            this MimeMessage original, string toAddress, string subjectSuffix = "", string fromAddress = "")
        {
            if (original == null)
                throw new ArgumentNullException(nameof(original));

            var mFrom = EmailHelper.ParseMailboxAddress(fromAddress);
            var mTo = EmailHelper.ParseMailboxAddress(toAddress);
            string subject = string.Join(' ', original.Subject ?? "", subjectSuffix);
            var mimeBody = original.Body;

            return new MimeMessage(mFrom, mTo, subject, mimeBody);
        }

        public static MimeMessage CreatePrefixedMessage(
            this MimeMessage original, string prefixSubject = null, string prefixBody = null,
            bool quoteOriginal = true, IEnumerable<MailboxAddress> mReplyTo = null)
        {
            if (original == null)
                return null;

            var mFrom = original.From.Mailboxes.FormatMailboxName();
            var mTo = original.To.Mailboxes.FormatMailboxName();
            if (mReplyTo.IsNullOrEmpty())
                mReplyTo = original.ReplyTo.Mailboxes;

            string subject = EmailHelper.GetPrefixedSubject(original.Subject, prefixSubject);

            var mimeBody = original.Body;
#if DEBUG
            //if (!string.IsNullOrEmpty(prefixBody))
            //{
            //    var visitor = new MimeMessageVisitor(prefixBody);
            //    visitor.Visit(original);

            //    if (visitor.Body != null)
            //        mimeBody = visitor.Body;

            //    //mimeBody = original.PrependText(prefixBody); //TODO test
            //}
#endif

            var message = new MimeMessage(mFrom, mTo, subject, mimeBody);
            message.ReplyTo.AddRange(mReplyTo);

            string logMessage = string.Format("Sending to: {0}, subject: '{1}'",
                mTo.Select(a => a.Address).ToEnumeratedString(), subject);
            if (_logger != null && _logger != default)
                _logger.LogDebug(logMessage);
            else
                Trace.TraceInformation(logMessage);

            return message;
        }

        //public static MimeEntity PrependText(this MimeMessage original, string prefix = "")
        //{
        //    var mimeBodyParts = new List<MimeEntity>();
        //    using (var iter = new MimeIterator(original))
        //    {
        //        while (iter.MoveNext())
        //        {
        //            var textPart = iter.Current as TextPart;

        //            if (textPart != null && !textPart.IsAttachment)
        //            {
        //                textPart.Text = prefix + textPart.Text;
        //                mimeBodyParts.Add(textPart);
        //            }
        //            else
        //            {
        //                mimeBodyParts.Add(iter.Current);
        //            }
        //        }
        //    }
        //    return BuildMultipart(mimeBodyParts);
        //}

        //public static MimeEntity BuildMultipart(
        //    this IList<MimeEntity> mimeEntities)
        //{
        //    var multipart = new Multipart("mixed");
        //    if (mimeEntities.Count > 1)
        //        multipart.AddRange(mimeEntities);
        //    else if (mimeEntities.Count == 1)
        //        multipart.Add(mimeEntities[0]);
        //    return multipart;
        //}

        // quote the original message text
        internal static MimeEntity BuildTextPart(this MimeMessage original,
            string prependText = "", bool includeEmbedded = true, bool includeAttachments = true)
        {
            MimeEntity mimeBody = new TextPart();
            if (original == null)
                return mimeBody;

            bool isHtml = original.HtmlBody != null;
            var format = isHtml ? TextFormat.Html : TextFormat.Plain;

            try
            {
                var visitor = new ReplyVisitor(includeEmbedded, includeAttachments);
                if (!string.IsNullOrEmpty(prependText))
                    visitor.PrependMessage = prependText;
                visitor.Visit(original);

                mimeBody = visitor.Body != null ? visitor.Body :
                    new TextPart(format)
                    {
                        Text = original.QuoteOriginal(prependText)
                    };

                if (includeAttachments && visitor.Attachments.IsNotNullOrEmpty())
                {
                    mimeBody = mimeBody.BuildMultipart(visitor.Attachments.ToArray());
                }
            }
            catch (Exception ex)
            { // Should never happen
                _logger.LogError(ex, "Failed to quote original message.");
                mimeBody = new TextPart(format)
                {
                    Text = original.QuoteOriginal(prependText)
                };
                if (includeAttachments && original.Attachments.Any()) //TODO test
                {
                    mimeBody = mimeBody.BuildMultipart(original.Attachments.ToArray());
                }
            }

            return mimeBody;
        }

        public static bool IsCircularReference(this MimeMessage message)
        {
            var mailboxToCcAddresses = message == null ? Array.Empty<MailboxAddress>() :
                Enumerable.Concat(message.To.Mailboxes, message.Cc.Mailboxes);
            return mailboxToCcAddresses.EmailAddressesIntersect(message.From.Mailboxes);
        }

        public static bool EmailAddressesIntersect(
            this IEnumerable<MailboxAddress> mailboxes1, IEnumerable<MailboxAddress> mailboxes2)
        {
            return mailboxes1?.Select(m => m.Address)?.Intersect(mailboxes2
                ?.Select(m => m.Address), StringComparer.OrdinalIgnoreCase)?.Any() ?? false;
        }

        public static StringWriter QuoteEnvelope(this Envelope envelope, bool isHtml = true, string message = "", string bodyText = "")
        {
            //var timestamp = envelope.Date.GetValueOrDefault().ToLocalTime();
            //string rfc882DateTime = DateUtils.FormatDate(timestamp);
            var text = new StringWriter();
            if (!isHtml)
            {
                text.WriteLine(message);
                text.WriteLine();
                text.WriteLine("-------- Original Message --------");
                text.WriteLine("From: {0}", envelope.From);
                text.WriteLine("Sent: {0}", envelope.Date);
                text.WriteLine("To: {0}", envelope.To);
                if (envelope.Cc.Mailboxes.IsNotNullOrEmpty())
                    text.WriteLine("Cc: {0}", envelope.Cc);
                text.WriteLine("Subject: {0}", envelope.Subject);
                text.WriteLine();
                if (!string.IsNullOrEmpty(bodyText))
                {
                    text.WriteLine(bodyText ?? "");
                    text.WriteLine();
                }
            }
            else
            {
                //text.WriteLine("<body>");
                text.WriteLine("<div>");
                text.WriteLine(message);
                text.WriteLine("</div><br />");
                text.WriteLine("<div style=\"border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm\">");
                text.WriteLine("<div><p><hr />");
                text.WriteLine("<b>From:</b> {0}<br />", envelope.From.FormatMailboxNameAddress());
                text.WriteLine("<b>Sent:</b> {0}<br />", envelope.Date);
                text.WriteLine("<b>To:</b> {0}<br />", envelope.To.FormatMailboxNameAddress());
                if (envelope.Cc.Mailboxes.IsNotNullOrEmpty())
                    text.WriteLine("<b>Cc:</b> {0}<br />", envelope.Cc.FormatMailboxNameAddress());
                text.WriteLine("<b>Subject:</b> {0}<br />", envelope.Subject);
                text.WriteLine("</p></div><br />");
                text.WriteLine(bodyText ?? "");
                text.WriteLine("</div>");
                //text.WriteLine("</body>");
            }

            return text;
        }

        public static MimeMessage CreateMimeMessage(
            this MimeEntity attachment, string to, string subject, string bodyText,
            bool isHtml = true, string from = "", string replyTo = "")
        {
            var attachments = new MimeEntity[] { attachment };
            return attachments.CreateMimeMessage(to, subject, bodyText, isHtml, from, replyTo);
        }

        public static MimeMessage CreateMimeMessage(
            this IEnumerable<MimeEntity> attachments, string to, string subject, string bodyText,
            bool isHtml = true, string from = "", string replyTo = "")
        {
            EmailHelper.CreateMimeEnvelope(from, to, bodyText, isHtml, replyTo,
                out IEnumerable<MailboxAddress> mFrom, out IEnumerable<MailboxAddress> mTo,
                out MimeEntity mBody, out IEnumerable<MailboxAddress> mReplyTo);

            return EmailHelper.CreateMimeMessage(mFrom, mTo, subject, mBody, mReplyTo, attachments);
        }
    }
}
