using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;

namespace iku_aras_plugin
{
    public class ikuMail
    {
        public string  smtpserver { get; set; }
        public string port { get; set; }
        public string MailContent { get; set; }
        public string MailSubject { get; set; }
        public string RecipientAddress { get; set; }
        public string mailFrom { get; set; }
        public string formUser { get; set; }
        public string userPassword { get; set; }

        public ikuMail() { }

        public ikuMail(string Ismtpserver, string Iport, string IMailContent , string IMailSubject, string IRecipientAddress, string ImailFrom, string IformUser, string IuserPassword)
        {
            this.smtpserver = Ismtpserver;
            this.port = Iport;
            this.MailContent = IMailContent;
            this.MailSubject = IMailSubject;
            this.RecipientAddress = IRecipientAddress;
            this.mailFrom = ImailFrom;
            this.formUser = IformUser;
            this.userPassword = IuserPassword;

        }

        public bool SendMailResult()
        {
            SmtpClient SmtpClient = new SmtpClient();
            SmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            SmtpClient.Host = smtpserver;
            SmtpClient.Credentials = new System.Net.NetworkCredential(mailFrom, userPassword);

            MailMessage MailMessage =new MailMessage(mailFrom, RecipientAddress);
            MailMessage.From = new MailAddress(mailFrom, formUser);
            MailMessage.Subject = MailSubject;
            MailMessage.Body = MailContent;
            MailMessage.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessage.IsBodyHtml = true;

            try
            {
                SmtpClient.Send(MailMessage);
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        public void SendMail()
        {
            SmtpClient SmtpClient = new SmtpClient();
            SmtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            SmtpClient.Host = smtpserver;
            SmtpClient.Credentials = new System.Net.NetworkCredential(mailFrom, userPassword);

            MailMessage MailMessage = new MailMessage(mailFrom, RecipientAddress);
            MailMessage.From = new MailAddress(mailFrom, formUser);
            MailMessage.Subject = MailSubject;
            MailMessage.Body = MailContent;
            MailMessage.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessage.IsBodyHtml = true;

            try
            {
                SmtpClient.Send(MailMessage);
              //return true;
            }
            catch (Exception ex)
            {
              //  return false;
            }
        }

    }
}
