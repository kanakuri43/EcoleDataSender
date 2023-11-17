using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Text;
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
                // Log出力設定
                Console.SetOut(sw); 
                Console.WriteLine(string.Format("{0} Starting the process...", DateTime.Now.ToString("HH:mm:ss")));

                var config = LoadConfig();

                ReceiveResponseEmail();

                if (OutputDirectoryEmptyCheck(config) == false)
                {
                    Console.WriteLine(string.Format("{0} Output directory is not empty.", DateTime.Now.ToString("HH:mm:ss")));
                    return;
                }

                var csvFilePath = QueryExecutionAndTsvSave(config);                
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
                SmtpPassword = doc.Root.Element("Email").Element("SmtpPassword").Value,
                SubjectString = doc.Root.Element("Email").Element("SubjectString").Value
            };
        }

        static void ReceiveResponseEmail()
        {

        }

        static bool OutputDirectoryEmptyCheck(dynamic config)
        {
            Console.WriteLine(string.Format("{0} Checking the output directory is empty...", DateTime.Now.ToString("HH:mm:ss")));

            if (Directory.Exists($"{config.OutputFolder}"))
            {
                string[] files = Directory.GetFiles($"{config.OutputFolder}");
                return (files.Length == 0);
            }
            else { return false; }
        }

        static string QueryExecutionAndTsvSave(dynamic config)
        {
            Console.WriteLine(string.Format("{0} Executing SQL Select statement and Save data to TSV...", DateTime.Now.ToString("HH:mm:ss")));

            string connectionString = config.ConnectionString;
            string query = config.Query;
            StringBuilder tsvContent = new StringBuilder();
            var filePath = $"{config.OutputFolder}{DateTime.Now:yyyyMMdd-HHmmss}.tsv";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    // ヘッダー行の追加
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        tsvContent.Append(reader.GetName(i) + "\t");
                    }
                    tsvContent.AppendLine();

                    // データ行の追加
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            // タブ、改行、キャリッジリターンをエスケープ
                            var fieldValue = reader[i].ToString().Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                            tsvContent.Append(fieldValue + "\t");
                        }
                        tsvContent.AppendLine();
                    }
                }
            }
            // TSVをファイルに保存
            File.WriteAllText(filePath, tsvContent.ToString());
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
                Subject = config.IdentifySubject,
                Body = "Please find attached the data file."
            };

            message.Attachments.Add(new Attachment(filePath));

            client.Send(message);

            Console.WriteLine(string.Format("{0} Email sent successfully.", DateTime.Now.ToString("HH:mm:ss")));
        }

    }

}