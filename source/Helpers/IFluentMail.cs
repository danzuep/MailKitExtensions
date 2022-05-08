using MailKit;
using MailKitExtensions.Models;
using MimeKit;

namespace MailKitExtensions
{
		public interface IFluentMail
		{

			//IFluentMail From(string emailAddress, string name = null);

			IFluentMail To(string emailAddress, string name = null);

			//IFluentMail To(IEnumerable<MailboxAddress> mailAddresses);

			//IFluentMail CC(string emailAddress, string name = null);

			//IFluentMail CC(IEnumerable<MailboxAddress> mailAddresses);

			//IFluentMail BCC(string emailAddress, string name = null);

			//IFluentMail BCC(IEnumerable<MailboxAddress> mailAddresses);

			//IFluentMail ReplyTo(string address, string name = null);

			IFluentMail Subject(string subject);

			IFluentMail Body(string body, bool isHtml = true);

			//IFluentMail Attach(Attachment attachment);

			//IFluentMail Attach(IEnumerable<Attachment> attachments);

			//IFluentMail Attach(string fileName, string contentType = null, string contentId = null);

			//bool Send(CancellationToken token = default);

			Task<bool> SendAsync(CancellationToken token = default);

			EmailDto Message { get; }
		}
}
