using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using MailKitExtensions;

namespace Zue.MailKitExtensions.UnitTesting
{
    public class MailKitSenderTests
    {
        [Fact]
        public void Test1()
        {
            var email = Email
                .From("test")
                .To("test")
                .Subject("Hi")
                .Body("~");
            Assert.Equal("Hi", email.Message.Subject);
        }

        [Fact]
        public async Task Test2()
        {
            var email = Email
                .From("test")
                .To("test")
                .Subject("Hi")
                .Body("~");
            Email.SmtpHost = "myHost";
            await email.SendAsync();
            Assert.Equal("~", email.Message.Body);
        }
    }
}