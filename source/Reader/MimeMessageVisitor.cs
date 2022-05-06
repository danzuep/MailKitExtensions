using MimeKit;
using MimeKit.Text;
using MimeKit.Tnef;
using Zue.Common;

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
    public class MimeMessageVisitor : MimeVisitor
    {
        private static ILog _logger = LogUtil.GetLogger<MimeMessageVisitor>();

        private const string _invalidTag = "fontdefang_\""; //common ArgumentException
        private readonly string _prependMessage;
        private readonly Stack<Multipart> _stack = new Stack<Multipart>();
        private MimeMessage _original;
        private MimeEntity _body;
        public MimeEntity Body { get => _body; }

        /// <summary>
        /// Creates a new ReplyVisitor.
        /// </summary>
        public MimeMessageVisitor(string prependMessage = "")
        {
            _prependMessage = prependMessage;
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
                else
                {
                    //part.Accept(this);
                    Push(part);
                }
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
            TextConverter converter;

            if (entity.IsHtml)
            {
                converter = new HtmlToHtml()
                {
                    HtmlTagCallback = HtmlTagCallback
                };
            }
            else if (entity.IsFlowed)
            {
                entity.Text = _prependMessage + Environment.NewLine + entity.Text;

                var flowed = new FlowedToHtml();

                if (entity.ContentType.Parameters.TryGetValue("delsp", out string delsp))
                    flowed.DeleteSpace = delsp.ToLowerInvariant() == "yes";

                converter = flowed;
            }
            else
            {
                entity.Text = _prependMessage + Environment.NewLine + entity.Text;

                converter = new TextToHtml();
            }

            var part = new TextPart(TextFormat.Html)
            {
                Text = converter.Convert(entity.Text)
            };

            Push(part);
        }

        void Push(MimeEntity entity)
        {
            var multipart = entity as Multipart;

            if (_body == null)
            {
                _body = entity;
            }
            else if (_stack.Count > 0)
            {
                var parent = _stack.Peek();
                parent.Add(entity);
            }
            //else
            //{
            //    var parent = new Multipart();
            //    parent.Add(entity);
            //}

            if (multipart != null)
                _stack.Push(multipart);
        }

        void Pop()
        {
            if (_stack.Count > 0)
                _stack.Pop();
        }

        //http://www.mimekit.net/docs/html/Working-With-Messages.htm
        //http://www.mimekit.net/docs/html/M_MimeKit_MimeIterator_MoveNext.htm
        public static IList<MimePart> GetAttachments(MimeMessage message)
        {
            var attachments = new List<MimePart>();
            var multiparts = new List<Multipart>();

            using (var iter = new MimeIterator(message))
            {
                // collect our list of attachments and their parent multiparts
                while (iter.MoveNext())
                {
                    var multipart = iter.Parent as Multipart;
                    var part = iter.Current as MimePart;

                    if (multipart != null && part != null && part.IsAttachment)
                    {
                        // keep track of each attachment's parent multipart
                        multiparts.Add(multipart);
                        attachments.Add(part);
                    }
                }
            }

            // now remove each attachment from its parent multipart...
            for (int i = 0; i < attachments.Count; i++)
                multiparts[i].Remove(attachments[i]);

            return attachments;
        }

        void HtmlTagCallback(HtmlTagContext ctx, HtmlWriter htmlWriter)
        {
            if (ctx.TagId == HtmlTagId.Body && !ctx.IsEmptyElementTag)
            {
                if (ctx.IsEndTag)
                {
                    // pass the </body> tag through to the output
                    ctx.WriteTag(htmlWriter, true);
                }
                else
                {
                    // pass the <body> tag through to the output
                    ctx.WriteTag(htmlWriter, true);

                    // prepend the HTML reply with something descriptive
                    if (!string.IsNullOrEmpty(_prependMessage))
                    {
                        htmlWriter.WriteText(_prependMessage);
                        htmlWriter.WriteEmptyElementTag(HtmlTagId.Br);
                    }

                    ctx.InvokeCallbackForEndTag = true;
                }
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
