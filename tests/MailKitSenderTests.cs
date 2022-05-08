using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using MailKitExtensions;
using MailKitExtensions.Sender;

namespace Zue.MailKitExtensions.UnitTesting
{
    public class MailKitSenderTests
    {
        private readonly EmailSender _email =
            new EmailSender("smtp.gmail.com");

        [Fact]
        public void Test1()
        {
            var email = _email
                .From("test")
                .To("test")
                .Subject("Hi")
                .Body("~");
            Assert.Equal("Hi", email.Message.Subject);
        }

        [Fact]
        public async Task Test2()
        {
            var email = _email
                .From("test")
                .To("test")
                .Subject("Hi")
                .Body("~");
            await email.SendAsync();
            Assert.Equal("~", email.Message.Body);
        }
    }
}