using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using CsvHelper;
using System.Globalization;


namespace TimeSheetRemainder
{
    public class RemainderMail
    {
        List<Dictionary<string, string>> records = new List<Dictionary<string, string>>();

        public RemainderMail()
        {
            getMemberDetail();
        }

        public string checkDay()
        {
            string day = System.DateTime.Now.DayOfWeek.ToString();
            return day;
        }

        public List<string> getToEmailAddress()
        {
            List<string> To = new List<string>();
            foreach (var record in records)
            {
                foreach (var column in record)
                {
                    if(column.Key.Equals("To"))
                    {
                        To.Add(column.Value);
                    }
                }
            }

            foreach(string to in To)
            {
                Console.WriteLine(to);
            }

            return To;
        }

        public string getFromEmailAddress()
        {
            string From = string.Empty;
            foreach (var record in records)
            {
                foreach (var column in record)
                {
                    if(column.Key.Equals("From") && column.Value != null && !column.Value.Equals(""))
                    {
                        From = column.Value;
                    }
                }
            }
            Console.WriteLine(From);
            return From;
        }

        public void getMemberDetail()
        {
            string filePath = @"D:\Documents\Vasanth-Training\TimeSheetRemainder\RecipientListFolder\RecipientsList.csv";
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {

                csv.Read();
                csv.ReadHeader();
                var headers = csv.HeaderRecord;

                while (csv.Read())
                {
                    var record = new Dictionary<string, string>();
                    foreach (var header in headers)
                    {
                        record[header] = csv.GetField(header);
                    }
                    records.Add(record);
                }

                // foreach (var record in records)
                // {
                //     foreach (var column in record)
                //     {
                //         Console.WriteLine($"{column.Key}: {column.Value}");
                //     }
                //     Console.WriteLine();
                // }
            }
        }

        public string emailTemplate()
        {
            string templatePath = @"D:\Documents\Vasanth-Training\TimeSheetRemainder\HTMLTemplate\RemainderEmailTemplate.html";
            string htmlBody = File.ReadAllText(templatePath);
            htmlBody = htmlBody.Replace("{{Name}}", "Vasanth");
            //Console.WriteLine(htmlBody);
            return htmlBody;
        }

        public bool sendMail()
        {
            string senderEmailId = getFromEmailAddress();
            string subject = "Email Subject";
            //string attachmentPath = @"C:\Users\vasanth.raja\Documents\LifeRenewalData.csv";
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(senderEmailId);
                    foreach(string tomail in getToEmailAddress()) 
                    {
                        mail.To.Add(tomail);
                    }
                    mail.Subject = subject;
                    mail.Body = emailTemplate();
                    
                    // System.Net.Mail.Attachment attachment;
                    // attachment = new System.Net.Mail.Attachment(attachmentPath);
                    // mail.Attachments.Add(attachment);
                    NetworkCredential networkCredential = new NetworkCredential("vasanth.raja@aspiresys.com", "Priya@aspire123");
                    SmtpClient smtp = new SmtpClient
                    {
                        //Host = "smtp.office365.com",
                        Host = "smtp-mail.outlook.com",
                        UseDefaultCredentials = false,
                        EnableSsl = true,
                        Credentials = networkCredential,
                        Port = 25,
                        DeliveryMethod = SmtpDeliveryMethod.Network
                    };

                    smtp.Send(mail);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }
    }
}