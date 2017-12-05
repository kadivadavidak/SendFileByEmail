using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace SendFileByEmail
{
    internal class Program
    {
        internal static void SendEmail(string toEmail, string emailSubject, string fileLocation)
        {
            var mailMessage = new MailMessage();
            var smtpClient = new SmtpClient(ConfigurationManager.AppSettings["SmtpServer"]);
            mailMessage.From = new MailAddress(ConfigurationManager.AppSettings["FromEmailAddress"]);
            mailMessage.To.Add(toEmail);
            mailMessage.Bcc.Add("david.kirschman@springmobile.com");
            mailMessage.Subject = emailSubject;
            if (emailSubject.Contains("["))
            {
                mailMessage.Subject =
                    $"{emailSubject.Remove(0, emailSubject.IndexOf("[", StringComparison.Ordinal) + 1).Remove(emailSubject.Remove(0, emailSubject.IndexOf("[", StringComparison.Ordinal) + 1).Length - 1)} Report";
                if (mailMessage.Subject.Contains("["))
                {
                    mailMessage.Subject =
                        $"{mailMessage.Subject.Remove(0, mailMessage.Subject.IndexOf("[", StringComparison.Ordinal) + 1)}";
                    if (mailMessage.Subject.Contains("["))
                        mailMessage.Subject =
                            $"{mailMessage.Subject.Remove(0, mailMessage.Subject.IndexOf("[", StringComparison.Ordinal) + 1)}";
                }
            }
            mailMessage.Body = "See attachement.";
            var attachment = new Attachment(fileLocation);
            mailMessage.Attachments.Add(attachment);
            var num1 = 587;
            smtpClient.Port = num1;
            var networkCredential = new NetworkCredential(ConfigurationManager.AppSettings["SmtpUserName"], ConfigurationManager.AppSettings["SmtpPassword"]);
            smtpClient.Credentials = networkCredential;
            var num2 = 1;
            smtpClient.EnableSsl = num2 != 0;
            var message = mailMessage;
            smtpClient.Send(message);
        }

        internal static SqlConnection Connect()
        {
            return new SqlConnection(
                $"Data Source={(object) ConfigurationManager.AppSettings["ServerName"]};Initial Catalog={(object) ConfigurationManager.AppSettings["DatabaseName"]};User ID={(object) ConfigurationManager.AppSettings["ServerUserName"]};Password={(object) ConfigurationManager.AppSettings["ServerPassword"]}");
        }

        public static void GenerateFile(string storedProcName, string fileLocation)
        {
            DataTable dataTable = new DataTable();
            using (SqlConnection connection = Connect())
            {
                using (SqlCommand selectCommand = new SqlCommand(storedProcName, connection))
                {
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(selectCommand))
                    {
                        selectCommand.CommandTimeout = 900000;
                        selectCommand.CommandType = CommandType.StoredProcedure;
                        sqlDataAdapter.Fill(dataTable);
                    }
                }
            }
            string str = ",";
            StringBuilder stringBuilder = new StringBuilder();
            for (int index = 0; index < dataTable.Columns.Count; ++index)
            {
                stringBuilder.Append(dataTable.Columns[index]);
                if (index < dataTable.Columns.Count - 1)
                    stringBuilder.Append(str);
            }
            stringBuilder.AppendLine();
            foreach (DataRow row in (InternalDataCollectionBase)dataTable.Rows)
            {
                for (int index = 0; index < dataTable.Columns.Count; ++index)
                {
                    stringBuilder.Append("\"");
                    stringBuilder.Append(row[index].ToString().Replace("\"", ""));
                    stringBuilder.Append("\"");
                    if (index < dataTable.Columns.Count - 1)
                        stringBuilder.Append(str);
                }
                stringBuilder.AppendLine();
            }
            System.IO.File.WriteAllText(fileLocation, stringBuilder.ToString());
        }

        private static void Main(string[] args)
        {
            if (args.Length != 3)
                Console.WriteLine("Requires only three parameters.\n1. Name of stored procedure containing data to be emailed.\n2. Path to local storage to save the file to be emailed.\n3. Email address of recipient.");
            string storedProcName = args[0];
            string fileLocation1 = args[1];
            string toEmail = args[2];
            GenerateFile(storedProcName, fileLocation1);
            string emailSubject = storedProcName;
            string fileLocation2 = fileLocation1;
            SendEmail(toEmail, emailSubject, fileLocation2);
        }
    }
}
