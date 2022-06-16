# MailKitExtensions

## Usage

### Sending Mail
`
var smtp = new EmailSender("mail.example.com");
var email = smtp
    .From("me@example.com")
    .To("you@example.com")
    .Subject("Hi")
    .Body("~");
await email.SendAsync();
`
### Receiving Mail
`
var imap = new EmailReceiver(ImapNetworkCredential, "mail.example.com", "INBOX");
var emails = await imap.MailFolder.GetMimeMessagesAsync();
`
### Mail Folder Watcher
`
var imap = new IdleClientReceiver(ImapNetworkCredential, MessageHandlerDelegate, "mail.example.com", "INBOX");
Task.Run(() => imap.MailFolder.RunAsync());
`
