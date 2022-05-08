using MimeKit;
using MimeKit.Text;
using Zue.Common;
using MailKitExtensions.Sender;
using MailKitExtensions.Writer;
using MailKitExtensions.Attachments;
using MailKitExtensions.Receiver;
using MailKitExtensions.Models;

namespace MailKitExtensions
{
    public static class EmailExtensions
    {
        public static IFluentMail From(
            this EmailSender emailSender, string address, string name = "")
        {
            var email = new Email(emailSender);
            return email.From(address, name);
        }
    }

    // inspired by https://github.com/lukencode/FluentEmail/blob/master/src/FluentEmail.Core/Email.cs
    public class Email : IFluentMail
    {
        public readonly EmailSender _sender;

        private MimeMessage _mimeMessage = new MimeMessage();
        public EmailDto Message { get => new EmailDto(_mimeMessage); }

        public Email(EmailSender sender)
        {
            if (sender == null)
                throw new ArgumentNullException(nameof(sender));
            _sender = sender;
        }

        public IFluentMail From(string from)
        {
            _mimeMessage.From.AddRange(EmailHelper.ParseMailboxAddress(from));
            return this;
        }

        public IFluentMail From(string address, string name)
        {
            _mimeMessage.From.Add(new MailboxAddress(name, address));
            return this;
        }

        public IFluentMail To(string to)
        {
            _mimeMessage.To.AddRange(EmailHelper.ParseMailboxAddress(to));
            return this;
        }

        public IFluentMail To(string address, string name)
        {
            _mimeMessage.To.Add(new MailboxAddress(name, address));
            return this;
        }

        public IFluentMail Subject(string subject)
        {
            _mimeMessage.Subject = subject ?? string.Empty;
            return this;
        }

        public IFluentMail Body(string bodyText, bool isHtml = true)
        {
            if (_mimeMessage.Body != null)
            {
                var builder = new BodyBuilder();
                if (isHtml)
                    builder.HtmlBody = bodyText;
                else if (_mimeMessage.HtmlBody != null)
                    builder.HtmlBody = _mimeMessage.HtmlBody;
                if (!isHtml)
                    builder.TextBody = bodyText;
                else if (_mimeMessage.TextBody != null)
                    builder.TextBody = _mimeMessage.TextBody;
                if (_mimeMessage.Attachments != null)
                {
                    var linkedResources = _mimeMessage.Attachments
                        .Where(attachment => !attachment.IsAttachment);
                    if (linkedResources.Any())
                        builder.LinkedResources.AddRange(linkedResources);
                    var attachments = _mimeMessage.Attachments
                        .Where(attachment => attachment.IsAttachment);
                    if (linkedResources.Any())
                        builder.Attachments.AddRange(attachments);
                }
                _mimeMessage.Body = builder.ToMessageBody();
            }
            else
            {
                var format = isHtml ? TextFormat.Html : TextFormat.Plain;
                _mimeMessage.Body = new TextPart(format) { Text = bodyText ?? "" };
            }
            return this;
        }

        public IFluentMail Attach(params string[] filePaths)
        {
            var attachments = filePaths.GetMimeEntitiesFromFilePaths();
            _mimeMessage.Body = _mimeMessage.Body.BuildMultipart(attachments);
            return this;
        }

        public async Task<bool> SendAsync(CancellationToken ct = default)
        {
            return await _sender.SendAsync(_mimeMessage, ct);
        }
    }
}
