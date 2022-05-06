using MimeKit;
using MailKit;
using MailKitExtensions.Receiver;
using MailKitExtensions.Attachments;
using Zue.Common;

namespace MailKitExtensions.Models
{
    public class EmailDto : IDisposable
    {
        private Lazy<MimeMessage> _mimeMessage = new Lazy<MimeMessage>(new MimeMessage());
        protected MimeMessage MimeMessage { get => _mimeMessage.Value; }

        private UniqueId? UniqueId;
        public uint FolderIndex { get => UniqueId?.Id ?? 0; }

        private string _folderName;
        public string FolderName { get => _folderName ?? ""; private set => _folderName = value; }

        public string MessageId { get => MimeMessage.MessageId ?? ""; }
        public DateTimeOffset Date { get => MimeMessage.Date; }
        public IEnumerable<MailboxAddress> From { get => MimeMessage.From.Mailboxes; }
        public IEnumerable<MailboxAddress> To { get => MimeMessage.To.Mailboxes; }
        public IEnumerable<MailboxAddress> Cc { get => MimeMessage.Cc.Mailboxes; }
        public IEnumerable<MailboxAddress> Bcc { get => MimeMessage.Bcc.Mailboxes; }
        public IEnumerable<MailboxAddress> ResentFrom { get => MimeMessage.ResentFrom.Mailboxes; }
        public IEnumerable<MimeEntity> Attachments { get => MimeMessage.Attachments; }
        public IEnumerable<string> AttachmentNames { get => Attachments.GetAttachmentNames(); }
        public int AttachmentCount { get => Attachments.Count(); }
        public string AttachmentNamesEnumerated { get => AttachmentNames.ToEnumeratedString(); }
        public string Subject { get => MimeMessage.Subject ?? ""; }
        public bool IsHtml { get => MimeMessage.HtmlBody is not null; }
        public string Body { get => MimeMessage.HtmlBody ?? MimeMessage.TextBody ?? ""; }
        public string BodyText { get => IsHtml ? ReceiverMimeMessages.DecodeHtmlBody(Body) : MimeMessage.TextBody ?? ""; }

        private EmailDto() { }

        public EmailDto(MimeMessage mimeMessage)
        {
            _mimeMessage = new Lazy<MimeMessage>(mimeMessage);
        }

        public EmailDto(IMessageSummary mail, EmailReceiverOptions receiverOptions, CancellationToken ct = default)
        {
            UniqueId = mail?.UniqueId;
            FolderName = mail?.Folder?.FullName ?? "";
            if (UniqueId.HasValue)
            {
                using var emailReceiver = new EmailReceiver(receiverOptions);
                _mimeMessage = new Lazy<MimeMessage>(GetMimeMessageAsync(emailReceiver, ct).Result);
            }
        }

        private async ValueTask<MimeMessage> GetMimeMessageAsync(
            EmailReceiver emailReceiver, CancellationToken ct = default)
        {
            return await emailReceiver.GetMimeMessageAsync(UniqueId!.Value, FolderName, ct: ct);
        }

        public async Task<UniqueId?> MoveToMailFolderAsync(EmailReceiver emailReceiver,
            string destinationFolder, CancellationToken ct = default)
        {
            //using var emailReceiver = new EmailReceiver(receiverOptions);
            var mailFolder = emailReceiver.OpenFolder(destinationFolder, false, false, ct);
            var result = await emailReceiver.MoveToMailFolderAsync(
                UniqueId.Value, mailFolder, destinationFolder, ct: ct);
            if (result.Item1 != null)
            {
                UniqueId = result.Item1.Value;
                FolderName = destinationFolder;
            }
            return result.Item1;
        }

        public override string ToString()
        {
            using var text = new StringWriter();
            text.Write("'{0}' {1}. ", FolderName, FolderIndex);
            //text.Write("{0} Attachment(s){1}. ", AttachmentCount, AttachmentNamesEnumerated);
            text.Write("Sent: {0}. Received: {1}. Subject: '{2}'.", Date, DateTime.Now, Subject);
            return text.ToString();
        }

        public void Dispose()
        {
            if (_mimeMessage?.IsValueCreated ?? false)
                _mimeMessage = null;
        }
    }
}
