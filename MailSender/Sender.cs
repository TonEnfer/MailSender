using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace MailSender
{
    class Sender
    {
        public static void SendMail(string smtpServer, string from, string fromName, string password,
     List<string> mailto, string subject, string message, List<string> attachFiles)
        {
            try
            {
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(from, fromName),
                    DeliveryNotificationOptions = (DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.OnFailure),
                    Sender = new MailAddress(from),
                    Subject = subject,
                    Body = message,
                    IsBodyHtml = true

                };
                foreach (var m in mailto)
                    mail.To.Add(new MailAddress(m));
                mail.Bcc.Add(new MailAddress(from, fromName));

                foreach (var att in attachFiles)
                    if (!string.IsNullOrEmpty(att))
                        mail.Attachments.Add(new Attachment(att));

                SmtpClient client = new SmtpClient
                {
                    Host = smtpServer,
                    Port = 587,
                    EnableSsl = false,
                    Credentials = new NetworkCredential(from.Split('@')[0], password),
                    DeliveryMethod = SmtpDeliveryMethod.Network

                };
                Console.WriteLine("Отправка письма {0}", Convert.ToString(mail.To));

                client.Send(mail);
                mail.Dispose();
            }
            catch (Exception e)
            {
                throw new Exception("Mail.Send: " + e.Message);
            }
        }
    }
}
