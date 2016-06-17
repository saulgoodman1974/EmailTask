using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Mail;
using System.Collections.Specialized;
using System.IO;

namespace EmailTask
{

    class Program
    {

        //Get application settings
        public static string PathToAttachments = ConfigurationManager.AppSettings.Get("PathToFiles");
        public static string PathToEmailList = ConfigurationManager.AppSettings.Get("PathToEmailList");
        public static string MailServerHostName = ConfigurationManager.AppSettings.Get("MailServerHostName");
        public static string MailPort = ConfigurationManager.AppSettings.Get("MailPort");
        public static string TimeoutSetting = ConfigurationManager.AppSettings.Get("TimeoutSetting");
        public static string MessageSubject = ConfigurationManager.AppSettings.Get("MessageSubject");
        public static string MessageBody = ConfigurationManager.AppSettings.Get("MessageBody");

        static void Main()
        {
            
            SendEmails();          

        }

        private static void SendEmails()
        {

            IEnumerable<string> addresses = File.ReadLines(PathToEmailList);

            foreach (string itm in addresses)
            {

                try
                {
                    string address = itm.Trim();
                    SendEmail(address);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in SendEmails(): {0}",
                                ex.ToString());
                }
                                
            }

        }

        private static void SendEmail(string emailto)
        {

            string to = "jane@contoso.com";
            string from = "ben@contoso.com";
            MailMessage message = new MailMessage(from, emailto);

            message.Subject = MessageSubject;
            message.Body = MessageBody;

            SmtpClient client = new SmtpClient(MailServerHostName);

            // Credentials are necessary if the server requires the client 
            // to authenticate before it will send e-mail on the client's behalf.
            client.UseDefaultCredentials = true;
            client.Port = Convert.ToInt32(MailPort);
            client.Timeout = Convert.ToInt32(TimeoutSetting);

            //Add all the attachments to the email
            AddAttachments(message);

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in SendEmail(): {0}",
                            ex.ToString());
            }

        }

        private static MailMessage AddAttachments(MailMessage msg)
        {

            //get an enumerable collection of the file paths
            IEnumerable<string> attachments = Directory.EnumerateFiles(PathToAttachments);

            //loop through the paths
            foreach (String itm in attachments)
            {

                try
                {

                    //add the attachments to the message
                    msg.Attachments.Add(new Attachment(itm));

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in AddAttachments(): {0}",
                                ex.ToString());
                }
            }

            return msg;

        }

    }
}
