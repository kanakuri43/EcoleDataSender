using System;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Xml.Linq;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace EcoleDataSender
{
    class Program
    {
        static void Main()
        {
            StreamWriter sw = new StreamWriter(string.Format(@"log/{0}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss"), true));
            try
            {
                // 初期設定
                Console.SetOut(sw); 
                Console.WriteLine($"{DateTime.Now:HH:mm:ss} Starting the process...");
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var config = LoadConfig();

                // 更新完了メール確認
                // 更新完了メールは、本文に受信側が更新完了したファイル名が入ってくる
                /*
                var notifiedFileName = ReceiveEmail(config);
                if (notifiedFileName != "")
                {
                    // 通知があったらoutputフォルダから削除
                    DeleteFile(config, notifiedFileName);
                }
                */

                // outputフォルダが空かチェック
                // 通知メールの有無にかかわらず、出力フォルダが空だったら処理を開始する
                if (OutputDirectoryEmptyCheck(config) == false)
                {
                    // 以前のファイルgア残っていたらファイル作成しないで終了
                    Console.WriteLine($"{DateTime.Now:HH:mm:ss} Output directory is not empty.");
                    return;
                }

                // 更新対象データでTSV作成
                //var tsvFilePath = QueryExecutionAndSaveTsv(config);
                // 更新対象データでSQLiteファイル作成
                var sqliteFilePath = QueryExecutionAndSaveSQLite(config);

                // メールで送信
                /*SendEmail(config, tsvFilePath);
                */
                Console.WriteLine($"{DateTime.Now:HH:mm:ss} Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now:HH:mm:ss} Error: {ex.Message}");
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
                Query           = doc.Root.Element("Database").Element("Query").Value,
                OutputFolder    = doc.Root.Element("Output").Element("Folder").Value,

                Pop3Server      = doc.Root.Element("Email").Element("Pop3").Element("Server").Value,
                Pop3Port        = int.Parse(doc.Root.Element("Email").Element("Pop3").Element("Port").Value),
                Pop3User        = doc.Root.Element("Email").Element("Pop3").Element("User").Value,
                Pop3Password    = doc.Root.Element("Email").Element("Pop3").Element("Password").Value,
                ReceiveSubject  = doc.Root.Element("Email").Element("Pop3").Element("Subject").Value,

                EmailTo         = doc.Root.Element("Email").Element("Smtp").Element("To").Value,
                EmailFrom       = doc.Root.Element("Email").Element("Smtp").Element("From").Value,
                SmtpServer      = doc.Root.Element("Email").Element("Smtp").Element("Server").Value,
                SmtpPort        = int.Parse(doc.Root.Element("Email").Element("Smtp").Element("Port").Value),
                SmtpUser        = doc.Root.Element("Email").Element("Smtp").Element("User").Value,
                SmtpPassword    = doc.Root.Element("Email").Element("Smtp").Element("Password").Value,
                
                SendSubject     = doc.Root.Element("Email").Element("Smtp").Element("Subject").Value + doc.Root.Element("Company").Element("Id").Value
            };
        }

        static string ReceiveEmail(dynamic config)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Fetching email attachment...");

            using var client = new Pop3Client();
            client.Connect(config.Pop3Server, config.Pop3Port, true);
            client.Authenticate(config.Pop3User, config.Pop3Password);

            var messageCount = client.GetMessageCount();

            // Find the first message with the specified subject and an attachment
            Message attachmentMessage = null;
            for (int i = messageCount; i >= 1; i--)
            {
                var message = client.GetMessage(i);
                if (message.Headers.Subject.Contains(config.ReceiveSubject))
                {                    
                    Console.WriteLine($"{DateTime.Now:HH:mm:ss} Updated notification mail received.");
                    var body = message.FindFirstPlainTextVersion().GetBodyAsText();
                    client.DeleteMessage(i);
                    client.Disconnect();
                    Console.WriteLine($"{DateTime.Now:HH:mm:ss} Updated notification mail deleted.");

                    return (body);
                }
            }
            
            return "";
        }

        static void DeleteFile(dynamic config, string fileName)
        {            
            Console.WriteLine($"{DateTime.Now:HH:mm:ss} Deleting files that have been updated.");

            if (File.Exists($"{config.OutputFolder}" + fileName))
            {
                File.Delete($"{config.OutputFolder}" + fileName);
            }
            else
            {
                // ファイルがなければ何もしない
                Console.WriteLine($"{DateTime.Now:HH:mm:ss} Target File Not Found.");
            }
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

        static string QueryExecutionAndSaveTsv(dynamic config)
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

        static string QueryExecutionAndSaveSQLite(dynamic config)
        {
            Console.WriteLine(string.Format("{0} Executing SQL Select statement and Save data to SQLIte DB...", DateTime.Now.ToString("HH:mm:ss")));

            string connectionString = config.ConnectionString;
            string query = config.Query;
            var sqliteFileName = $"{config.OutputFolder}{DateTime.Now:yyyyMMdd}.sqlite";

            using (SqlConnection sqlCon = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, sqlCon);
                sqlCon.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    var sqliteCon = new SQLiteConnection($"Data Source={sqliteFileName};Version=3;");
                    sqliteCon.Open();
                    using (var command = new SQLiteCommand(sqliteCon))
                    {
                        var sql = "CREATE TABLE updated_items ( "
                                + "エコールコード INTEGER"
                                + ", 商品名 TEXT "
                                + ", 品番 TEXT "
                                + ", 分類コード TEXT "
                                + ", 単位 TEXT "
                                + ", 表示定価 REAL "
                                + ", 商品メーカー名 TEXT "
                                + ")";
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                    }

                    // データ行の追加
                    while (reader.Read())
                    {
                        using (var command = new SQLiteCommand(sqliteCon))
                        {
                            var sql = "INSERT INTO updated_items "
                                    + "VALUES ("
                                    + (int)reader["エコールコード"]
                                    + ", '" + reader["商品名"].ToString() + "'"
                                    + ", '" + reader["品番"].ToString() + "'"
                                    + ", '" + reader["分類コード"].ToString() + "'"
                                    + ", '" + reader["単位"].ToString() + "'"
                                    + ", '" + Convert.ToSingle(reader["表示定価"]) + "'"
                                    + ", '" + reader["商品メーカー名"].ToString() + "'"
                                    + ")";
                            command.CommandText = sql;
                            command.ExecuteNonQuery();
                        }

                    }
                }
            }

            return sqliteFileName;
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
                Subject = config.SendSubject,
                Body = ""
            };

            message.Attachments.Add(new Attachment(filePath));

            client.Send(message);

            Console.WriteLine(string.Format("{0} Email sent successfully.", DateTime.Now.ToString("HH:mm:ss")));
        }

    }

}