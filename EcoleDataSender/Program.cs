using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Xml.Linq;

namespace EcoleDataSender
{
    class Program
    {
        static void Main()
        {
            StreamWriter sw = new StreamWriter(string.Format(@"log/{0}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss"), true));
            try
            {
                Console.SetOut(sw); // 出力先を設定
                Console.WriteLine(string.Format("{0} Starting the process...", DateTime.Now.ToString("HH:mm:ss")));

                var config = LoadConfig();

                var dataTable = ExecuteSelect(config);
                var csvFilePath = SaveToCsv(dataTable, config.OutputFolder);

                SendEmail(config, csvFilePath);

                Console.WriteLine(string.Format("{0} Process completed successfully.", DateTime.Now.ToString("HH:mm:ss")));
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0} Error: {1}", DateTime.Now.ToString("HH:mm:ss"), ex.Message));
            }
            finally
            {
                sw.Dispose();
            }
        }

        static dynamic LoadConfig()
        {
            var doc = XDocument.Load("config.xml");

            return new
            {
                ConnectionString = doc.Root.Element("Database").Element("ConnectionString").Value,
                Query = doc.Root.Element("Database").Element("Query").Value,
                OutputFolder = doc.Root.Element("Output").Element("Folder").Value,
                EmailTo = doc.Root.Element("Email").Element("To").Value,
                EmailFrom = doc.Root.Element("Email").Element("From").Value,
                SmtpServer = doc.Root.Element("Email").Element("SmtpServer").Value,
                SmtpPort = int.Parse(doc.Root.Element("Email").Element("SmtpPort").Value),
                SmtpUser = doc.Root.Element("Email").Element("SmtpUser").Value,
                SmtpPassword = doc.Root.Element("Email").Element("SmtpPassword").Value
            };
        }

        static System.Data.DataTable ExecuteSelect(dynamic config)
        {
            Console.WriteLine(string.Format("{0} Executing SQL Select statement...", DateTime.Now.ToString("HH:mm:ss")));

            using var connection = new SqlConnection(config.ConnectionString);
            connection.Open();

            using var command = new SqlCommand(config.Query, connection);

            var reader = command.ExecuteReader();
            var dataTable = new System.Data.DataTable();
            dataTable.Load(reader);

            Console.WriteLine(string.Format("{0} SQL Select statement executed successfully.", DateTime.Now.ToString("HH:mm:ss")));
            return dataTable;
        }

        static string SaveToCsv(System.Data.DataTable table, string folder)
        {
            Console.WriteLine(string.Format("{0} Saving data to CSV...", DateTime.Now.ToString("HH:mm:ss")));

            var filePath = $"{folder}{DateTime.Now:yyyyMMdd-HHmmss}.csv";
            using var writer = new StreamWriter(filePath);
            foreach (System.Data.DataRow row in table.Rows)
            {
                var items = row.ItemArray;
                writer.WriteLine(string.Join(",", items));
            }

            Console.WriteLine(string.Format("{0} Data saved to {1}", DateTime.Now.ToString("HH:mm:ss"), filePath));
            return filePath;
        }

        static void SendEmail(dynamic config, string filePath)
        {
            Console.WriteLine(string.Format("{0} Sending email...", DateTime.Now.ToString("HH:mm:ss")));

            using var client = new SmtpClient(config.SmtpServer, config.SmtpPort)
            {
                Credentials = new System.Net.NetworkCredential(config.SmtpUser, config.SmtpPassword),
                EnableSsl = true
            };

            using var message = new MailMessage(config.EmailFrom, config.EmailTo)
            {
                Subject = "エコール商品マスタ",
                Body = "Please find attached the data file."
            };

            message.Attachments.Add(new Attachment(filePath));

            client.Send(message);

            Console.WriteLine(string.Format("{0} Email sent successfully.", DateTime.Now.ToString("HH:mm:ss")));
        }

    }

}