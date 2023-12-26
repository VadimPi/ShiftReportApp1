using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using System.IO;
using NLog;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace ShiftReportApp1
{
    internal class Sending
    {
        static string LoginMail {  get; set; }
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public static void SendEmail(string toAddress, string subject, string body, byte[] attachmentBytes = null)
        {
            try
            {
                using (SmtpClient smtpClient = ConfigureSmtpClient())
                {
                    using (MailMessage mailMessage = CreateMailMessage(toAddress, subject, body, attachmentBytes))
                    {
                        smtpClient.Send(mailMessage);
                    }

                    Logger.Info("Письмо успешно отправлено!");
                    MessageBox.Show("Письмо успешно отправлено!", "info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка при отправке письма: {ex.Message}");
                // Обработка ошибок
            }
        }

        private static SmtpClient ConfigureSmtpClient() // Конфигурирование сервера
        {
            string PasswordMail = "";
            string filePath = "CS.tx_";
            if (File.Exists(filePath))
            {
                string[] settings = File.ReadAllText(filePath).Split(',');

                if (settings.Length == 7)
                {
                    LoginMail = settings[5];
                    PasswordMail = settings[6];
                }
                else
                {
                    Console.WriteLine("Invalid format in the settings file.");
                }
            }
            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(LoginMail, PasswordMail),
                EnableSsl = true
            };

            return smtpClient;
        }

        private static MailMessage CreateMailMessage(string toAddress, string subject, string body, byte[] attachmentBytes) // Создание письма
        {
            MailMessage mailMessage = new MailMessage
            {
                From = new MailAddress(LoginMail),
                To = { toAddress },
                Subject = subject,
                Body = body
            };

            if (attachmentBytes != null)
            {
                MemoryStream memoryStream = new MemoryStream(attachmentBytes);
                Attachment attachment = new Attachment(memoryStream, "имя_файла.pdf");
                mailMessage.Attachments.Add(attachment);
            }

            return mailMessage;
        }
    }
}
