using MailKit.Net.Smtp;
using MimeKit;

namespace Services
{
    public class EmailService
    {
        public static async Task SendEmailAsync(string email, string subject, string message)
        {
            var emailMessage = new MimeMessage();
            var login = Environment.GetEnvironmentVariable("EmailLogin");
            var password = Environment.GetEnvironmentVariable("EmailPassword");
            emailMessage.From.Add(new MailboxAddress("Teams Meeting Migration Service", login));
            emailMessage.To.Add(new MailboxAddress("", email));
            emailMessage.Subject = subject;
            emailMessage.Body = new TextPart(MimeKit.Text.TextFormat.Plain)
            {
                Text = message
            };

            using var client = new SmtpClient();
            await client.ConnectAsync("smtp.yandex.ru", 465, true);
            await client.AuthenticateAsync(login, password);
            await client.SendAsync(emailMessage);
 
            await client.DisconnectAsync(true);
        }
    }
}
