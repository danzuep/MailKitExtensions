using System.Net;

namespace MailKitExtensions.Models
{
    public record EmailOptions
    {
        public EmailSenderOptions Sender { get; set; }
        public EmailReceiverOptions Receiver { get; set; }

        //public EmailOptions DecryptSettings()
        //{
        //    if (Receiver?.ImapCredential is NetworkCredential imapCredential)
        //        Receiver.ImapCredential = CryptographyHelper.Decrypt(imapCredential);
        //    if (Sender?.SmtpCredential is NetworkCredential smtpCredential)
        //        Sender.SmtpCredential = CryptographyHelper.Decrypt(smtpCredential);
        //    return this;
        //}
    }

    public record EmailSenderOptions
    {
        public NetworkCredential? SmtpCredential { get; set; }
        public string SmtpHost { get; set; } // server
        public string FromAddress { get; set; } // noreply@example.com
        public string ReplyToAddress { get; set; }
        public string NoReplyAddress { get; set; }
        public string ToAddress { get; set; }
        public string Subject { get; set; } = "Do Not Reply";
        public string BodyHtml { get; set; } = "<body>{{BodyText}}</body>{{Signature}}";
        public string BodyText { get; set; } = "";
        public string Signature { get; set; } = "";
        public string AttachmentFilePath { get; set; } = "";
    }

    public record EmailReceiverOptions
    {
        public string FolderToProcess { get; set; } = "INBOX";
        public string ImapHost { get; set; }
        public NetworkCredential ImapCredential { get; set; }
        public string DownloadPath { get; set; } = "C:\\Temp\\Emails";
    }
}
