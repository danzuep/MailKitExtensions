using MimeKit;
using MimeKit.Text;
using MimeKit.Tnef;
using MimeKit.Utils;
using Microsoft.Extensions.Logging;
using Zue.Common;
using MailKitExtensions.Helpers;
using MailKitExtensions.Attachments;

/* An email might have the following structure:
 * multipart/mixed
   multipart/alternative
      text/plain
      multipart/related
         text/html
         image/jpeg
         image/png
   application/octet-stream
   application/zip
 */

namespace MailKitExtensions.Reader
{
    //http://www.mimekit.net/docs/html/Working-With-Messages.htm
    //http://www.mimekit.net/docs/html/Frequently-Asked-Questions.htm#Reply
    //http://www.mimekit.net/docs/html/Frequently-Asked-Questions.htm#MessageBody
    public class ReplyVisitor : MimeVisitor
    {
        private static ILogger _logger = LogUtil.GetLogger<ReplyVisitor>();

        public MimeEntity Body { get => _body; }
        public IList<MimeEntity> Attachments { get; private set; } = new List<MimeEntity>();
        public string BorderCssLeft { get; set; } = "border-left: 1px solid #ccc; margin: 0 0 0 0; padding: 0 0 0 .8em;";
        public string BorderCssBottom { get; set; } = "border-bottom: 1px solid #E1E1E1; margin: 0 0 1em 0; padding: 0 0 0.8em 0;";
        public string PrependMessage { get; set; } = string.Empty;
        public string SectionBreakMessage { get; set; } = "-----Original Message-----";

        readonly Stack<Multipart> _stack = new Stack<Multipart>();
        MimeMessage _original;
        MimeEntity _body;
        bool _includeEmbedded;
        bool _includeAttachments;

        /// <summary>
        /// Creates a new ReplyVisitor.
        /// </summary>
        public ReplyVisitor(bool includeEmbedded = true, bool includeAttachments = true)
        {
            _includeEmbedded = includeEmbedded;
            _includeAttachments = includeAttachments;
        }

        /// <summary>
        /// Visit the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        public override void Visit(MimeMessage message)
        {
            _original = message;
            _stack.Clear();

            base.Visit(message);
        }

        protected override void VisitMultipart(Multipart multipart)
        {
            foreach (var part in multipart)
            {
                if (part is MultipartAlternative)
                    part.Accept(this);
                else if (part is MultipartRelated)
                    part.Accept(this);
                else if (part is TextPart)
                    part.Accept(this);
                else if (part is MimePart)
                    part.Accept(this);
                else if (part is MessagePart)
                    part.Accept(this);
                else if (part is TnefPart)
                    part.Accept(this);
                else if (_includeAttachments)
                    Attachments.Add(part);
            }
        }

        protected override void VisitMultipartAlternative(MultipartAlternative alternative)
        {
            var multipart = new MultipartAlternative();

            // push this multipart/alternative onto our stack
            Push(multipart);

            // visit each of the multipart/alternative children
            for (int i = 0; i < alternative.Count; i++)
                alternative[i].Accept(this);

            // pop this multipart/alternative off our stack
            Pop();
        }

        protected override void VisitMultipartRelated(MultipartRelated related)
        {
            var multipart = new MultipartRelated();
            var root = related.Root;

            // push this multipart/related onto our stack
            Push(multipart);

            // visit the root document
            root.Accept(this);

            // visit each node
            for (int i = 0; i < related.Count; i++)
            {
                if (related[i] != root)
                    related[i].Accept(this);
            }

            // pop this multipart/related off our stack
            Pop();
        }

        protected override void VisitTextPart(TextPart entity)
        {
            if (!entity.IsAttachment)
            {
                TextConverter converter;

                if (entity.IsHtml)
                {
                    converter = new HtmlToHtml
                    {
                        HtmlTagCallback = HtmlTagCallback
                    };
                }
                else if (entity.IsFlowed)
                {
                    var flowed = new FlowedToHtml();

                    if (entity.ContentType.Parameters.TryGetValue("delsp", out string delsp))
                        flowed.DeleteSpace = delsp.ToLowerInvariant() == "yes";

                    converter = flowed;
                }
                else
                {
                    converter = new TextToHtml();
                    entity.Text = QuoteOriginal();
                }

                var part = new TextPart(TextFormat.Html)
                {
                    Text = converter.Convert(entity.Text)
                };

                Push(part);
            }
            else if (_includeAttachments)
            {
                Attachments.Add(entity);
            }
        }

        protected override void VisitTnefPart(TnefPart entity)
        {
            // extract any attachments in the MS-TNEF part
            if (_includeAttachments)
                Attachments.AddRange(entity.ExtractAttachments());
        }

        protected override void VisitMessagePart(MessagePart entity)
        {
            // treat message/rfc822 parts as attachments
            if (_includeAttachments)
                Attachments.Add(entity);
        }

        protected override void VisitMimePart(MimePart entity)
        {
            if (_includeAttachments && entity.IsAttachment)
                Attachments.Add(entity);
            else if (_includeEmbedded && !entity.IsAttachment)
                Push(entity);
        }

        void Push(MimeEntity entity)
        {
            var multipart = entity as Multipart;

            if (!entity.IsAttachment)
            {
                if (_body == null)
                {
                    _body = entity;
                }
                else if (_stack.Count > 0)
                {
                    var parent = _stack.Peek();
                    parent.Add(entity);
                }
                else
                {
                    _logger.LogWarning("Unknown MimeEntity pushed to the body stack, adding as attachment instead. Message ID: {0}", _original.MessageId);
                    Attachments.Add(entity);
                }
            }
            else if (_includeAttachments)
            {
                if (_stack.Count > 0)
                {
                    var parent = _stack.Peek();
                    parent.Add(entity);
                }
                else if (entity is TextPart)
                {
                    _logger.LogInformation("Extra MimeEntity TextPart attachment tried to push to the body stack. Message ID: {0}", _original.MessageId);
                }
                _logger.LogInformation("{0} attachment added as Multipart.", entity.GetType().Name);
                Attachments.Add(entity);
            }

            if (multipart != null)
                _stack.Push(multipart);
        }

        void Pop()
        {
            _stack.Pop();
        }

        private string QuoteOriginal()
        {
            var original = _original;
            //var timestamp = envelope.Date.GetValueOrDefault().ToLocalTime();
            string rfc882DateTime = DateUtils.FormatDate(original.Date);
            using var text = new StringWriter();
            text.WriteLine(PrependMessage);
            text.WriteLine();
            text.WriteLine(SectionBreakMessage);
            text.WriteLine("From: {0}", original.From.FormatMailboxNameAddress());
            text.WriteLine("Sent: {0}", rfc882DateTime);
            text.WriteLine("To: {0}", original.To.FormatMailboxNameAddress());
            if (!original.Cc.Mailboxes.IsNullOrEmpty())
                text.WriteLine("Cc: {0}", original.Cc.FormatMailboxNameAddress());
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

            return text.ToString();
        }

        void FormatMessageEnvelope(HtmlTagContext ctx, HtmlWriter htmlWriter)
        {
            if (ctx.IsEndTag)
            {
                // end our opening <blockquote>
                htmlWriter.WriteEndTag(HtmlTagId.BlockQuote);

                // pass the </body> tag through to the output
                ctx.WriteTag(htmlWriter, true);
            }
            else
            {
                // pass the <body> tag through to the output
                ctx.WriteTag(htmlWriter, true);

                // prepend the HTML reply with something descriptive
                if (!string.IsNullOrEmpty(PrependMessage))
                {
                    htmlWriter.WriteText(PrependMessage);
                    htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                    htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                }

                htmlWriter.WriteStartTag(HtmlTagId.P);
                htmlWriter.WriteAttribute(HtmlAttributeId.Style, BorderCssBottom);
                htmlWriter.WriteText(SectionBreakMessage);
                htmlWriter.WriteEndTag(HtmlTagId.P);

                // Wrap the original content in a <blockquote>
                htmlWriter.WriteStartTag(HtmlTagId.BlockQuote);
                htmlWriter.WriteAttribute(HtmlAttributeId.Style, BorderCssLeft);

                htmlWriter.WriteStartTag(HtmlTagId.B);
                htmlWriter.WriteText("From: ");
                htmlWriter.WriteEndTag(HtmlTagId.B);
                htmlWriter.WriteText(_original.From.FormatMailboxNameAddress());
                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                htmlWriter.WriteStartTag(HtmlTagId.B);
                htmlWriter.WriteText("Sent: ");
                htmlWriter.WriteEndTag(HtmlTagId.B);
                string rfc882DateTime = DateUtils.FormatDate(_original.Date);
                htmlWriter.WriteText(rfc882DateTime);
                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                htmlWriter.WriteStartTag(HtmlTagId.B);
                htmlWriter.WriteText("To: ");
                htmlWriter.WriteEndTag(HtmlTagId.B);
                htmlWriter.WriteText(_original.To.FormatMailboxNameAddress());
                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                if (_original.Cc.Mailboxes.IsNotNullOrEmpty())
                {
                    htmlWriter.WriteStartTag(HtmlTagId.B);
                    htmlWriter.WriteText("CC: ");
                    htmlWriter.WriteEndTag(HtmlTagId.B);
                    htmlWriter.WriteText(_original.Cc.FormatMailboxNameAddress());
                    htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                }

                htmlWriter.WriteStartTag(HtmlTagId.B);
                htmlWriter.WriteText("Subject: ");
                htmlWriter.WriteEndTag(HtmlTagId.B);
                htmlWriter.WriteText(_original.Subject);
                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                if (_original.Attachments.IsNotNullOrEmpty())
                {
                    htmlWriter.WriteStartTag(HtmlTagId.B);
                    htmlWriter.WriteText(string.Format(
                        "Attachment{0}: ", _original.Attachments.Count().S()));
                    htmlWriter.WriteEndTag(HtmlTagId.B);
                    htmlWriter.WriteText(_original.Attachments
                        .GetAttachmentNames().ToEnumeratedString());
                    htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                }

                htmlWriter.WriteStartTag(HtmlTagId.B);
                htmlWriter.WriteText("ID: ");
                htmlWriter.WriteEndTag(HtmlTagId.B);
                htmlWriter.WriteText(_original.MessageId);
                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                if (_original.ResentFrom.Mailboxes.IsNotNullOrEmpty())
                {
                    htmlWriter.WriteStartTag(HtmlTagId.B);
                    htmlWriter.WriteText("Resent From: ");
                    htmlWriter.WriteEndTag(HtmlTagId.B);
                    htmlWriter.WriteText(_original.ResentFrom.FormatMailboxNameAddress());
                    htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                }

                htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);

                ctx.InvokeCallbackForEndTag = true;
            }
        }

        void HtmlTagCallback(HtmlTagContext ctx, HtmlWriter htmlWriter)
        {
            const string _invalidTag = "fontdefang_\""; //common ArgumentException

            if (ctx.TagId == HtmlTagId.Body && !ctx.IsEmptyElementTag)
            {
                FormatMessageEnvelope(ctx, htmlWriter);
            }
            else if (ctx.TagName != _invalidTag)
            {
                try
                {
                    // pass the tag through to the output
                    ctx.WriteTag(htmlWriter, true);
                }
                catch (ArgumentException ex)
                {
                    _logger.LogWarning("Tag name: '{0}'. Attributes: '{1}'. Message ID: {2}. {3}",
                        ctx.TagName, ctx.Attributes?.Select(a => a.Name).ToEnumeratedString(), _original.MessageId, ex.Message);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to parse email HTML tags. Message ID: {0}", _original.MessageId);
                    System.Diagnostics.Debugger.Break();
                }
            }
            else
            {
                _logger.LogInformation("Invalid HTML tag not processed: '{0}'. Message ID: {1}", _invalidTag, _original.MessageId);
            }
        }
    }
}
