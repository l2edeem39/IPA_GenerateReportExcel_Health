using GemBox.Spreadsheet;
using IPA_GenerateReportExcel_Health.Models;
using Renci.SshNet;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

namespace IPA_GenerateReportExcel_Health
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileLog = string.Format(".//Log//Health_{0:ddMMyyyy}_log.txt", DateTime.Now);
            WriteLogSystem("Start Process =>", fileLog);
            try
            {
                var fileName = string.Format("Health_{0:ddMMyyyy_hhmm}.xlsx", DateTime.Now);
                var conn = new SqlConnection(Configulation.Db_r4ad01.ToString().Trim());

                //Open Connection
                conn.Open();
                var guidId = Guid.NewGuid().ToString();
                WriteLogSystem("==> "+ guidId + " DB Connected", fileLog);
                var config = new ConfigModel();

                using (SqlCommand command = new SqlCommand("SpIPAGetConfig", conn))
                {
                    command.Parameters.AddWithValue("@subject", Configulation.Type);
                    command.CommandType = System.Data.CommandType.StoredProcedure;

                    DataTable dtExcel;
                    DataTable dtConfig;

                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataSet dataSet = new DataSet();
                        adapter.Fill(dataSet);
                        dtConfig = dataSet.Tables[0];
                        dtExcel = dataSet.Tables[1];
                    }

                    var resultGen = GenerateExcel(dtExcel, config.ExcelFilePath + fileName);

                    config = MappingDatatableToMoel(dtConfig);

                    if (resultGen)
                    {
                        WriteLogSystem("==> Generate File Excel Completed.", fileLog);

                        #region Upload SFTP
                        var SFTPConfig = new SFTPConfigModel();
                        SFTPConfig.Host = config.UplaodHost;
                        SFTPConfig.Port = config.UplaodPort;
                        SFTPConfig.Username = config.UplaodUsername;
                        SFTPConfig.Password = config.UplaodPassword ;
                        SFTPConfig.Status = config.UplaodStatus;
                        SFTPConfig.FileDirectory = config.UplaodDirectory;
                        SFTPConfig.FileToUpload = config.UplaodFilePath;

                        var resultUplod = Upload(SFTPConfig);

                        if (resultUplod)
                        {
                            WriteLogSystem("==> Send SFTP File Completed.", fileLog);
                        }
                        else
                        {
                            WriteLogSystem("==> Config Not Send SFTP File.", fileLog);
                        }
                        #endregion

                        #region SendMail
                        
                        if (config.MailStatus == "1")
                        {
                            //Send Mail FTP Sucessful
                            var mailConfig = new MailConfigModel();
                            mailConfig.MailFrom = config.MailFrom;
                            mailConfig.MailTo = config.MailTo;
                            mailConfig.MailCc = config.MailCc;
                            mailConfig.Subject = config.MailSubject;
                            mailConfig.Body = config.MailBody;
                            mailConfig.Attachment = config.ExcelFilePath + fileName;
                            mailConfig.Smtp = config.MailSmtp;
                            var resultMail = SendMail(mailConfig);
                            if (resultMail)
                            {
                                WriteLogSystem("==> Send Mail Completed.", fileLog);
                            }
                        }
                        else
                        {
                            WriteLogSystem("==> Config Not Send Mail.", fileLog);
                        }
                        #endregion
                    }
                    else
                    {
                        WriteLogSystem("==> Generate File Excel Fail.", fileLog);
                    }

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                WriteLogSystem(ex.Message, fileLog);
            }
        }
        public static bool GenerateExcel(DataTable dt, string filePath)
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                var workbook = new ExcelFile();
                ExcelWorksheet worksheet = workbook.Worksheets.Add("Sheet 1");
                worksheet.InsertDataTable(dt,
                        new InsertDataTableOptions()
                        {
                            ColumnHeaders = true
                        });

                workbook.Save(".//FileExcel//DataExcel.xlsx");

                var workbook_last = ExcelFile.Load(".//FileExcel//DataExcel.xlsx");

                var worksheet_last = workbook_last.Worksheets.ActiveWorksheet;

                int columnCount = worksheet_last.CalculateMaxUsedColumns();
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet_last.Columns[i].AutoFit(1, worksheet_last.Rows[1], worksheet_last.Rows[worksheet_last.Rows.Count - 1]);
                }

                workbook_last.Save(filePath);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        class Configulation
        {
            public static string Db_r4ad01 = ConfigurationManager.AppSettings["Db_r4ad01"].ToString();
            public static string Type = ConfigurationManager.AppSettings["Type"].ToString();
        };
        public static bool SendMail(MailConfigModel mailConfig)
        {
            MailAddress to = new MailAddress(mailConfig.MailTo);
            MailAddress from = new MailAddress(mailConfig.MailFrom);
            Attachment data = new Attachment(mailConfig.Attachment, MediaTypeNames.Application.Octet);
            ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = System.IO.File.GetCreationTime(mailConfig.Attachment);
            disposition.ModificationDate = System.IO.File.GetLastWriteTime(mailConfig.Attachment);
            disposition.ReadDate = System.IO.File.GetLastAccessTime(mailConfig.Attachment);

            MailMessage email = new MailMessage(from, to);
            email.CC.Add(new MailAddress(mailConfig.MailCc));
            email.Subject = mailConfig.Subject;
            email.Body = mailConfig.Body;
            email.Attachments.Add(data);

            SmtpClient smtp = new SmtpClient(mailConfig.Smtp);
            smtp.UseDefaultCredentials = true;
            try
            {
                smtp.Send(email);
                return true;
            }
            catch (SmtpException ex)
            {
                throw ex;
            }
        }
        public static bool Upload(SFTPConfigModel SFTPConfig)
        {
            try
            {
                if (SFTPConfig.Status == "1")
                {
                    using (var sftpClient = new SftpClient(SFTPConfig.Host, Int32.Parse(SFTPConfig.Port), SFTPConfig.Username, SFTPConfig.Password))
                    using (var fs = new FileStream(SFTPConfig.FileToUpload, FileMode.Open))
                    {
                        sftpClient.Connect();
                        sftpClient.ChangeDirectory(SFTPConfig.FileDirectory);
                        //sftpClient.UploadFile(
                        //    fs,
                        //    "/ftproot/" + Path.GetFileName(fileToUpload),
                        //    uploaded =>
                        //    {
                        //        Console.WriteLine($"Uploaded {(double)uploaded / fs.Length * 100}% of the file.");
                        //    });

                        sftpClient.Disconnect();
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void WriteLogSystem(string logMessage, string PathfileLog)
        {
            if (logMessage == "Start Process =>")
            {
                using (StreamWriter w = File.CreateText(PathfileLog))
                {
                    w.WriteLine($"###############################################");
                    w.WriteLine($"{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")}");
                    w.WriteLine($"{logMessage}");
                }
            }
            else
            {
                using (StreamWriter w = File.AppendText(PathfileLog))
                {
                    w.WriteLine($"{DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss ")}" + $"{logMessage}");
                }
            }
        }
        public static void WriteLogHeader(string logMessage, string PathfileLog)
        {
            
        }
        public static void WriteLogDetail(string logMessage, string PathfileLog)
        {

        }
        public static ConfigModel MappingDatatableToMoel(DataTable dt)
        {
            var result = new ConfigModel();
            if (dt.Rows.Count > 0)
            {
                DataRow rowHead = dt.Rows[0];
                result.ExcelFilePath = rowHead["ExcelFilePath"] != null ? (string)rowHead["ExcelFilePath"].ToString() : string.Empty;
                result.LogFilePath = rowHead["LogFilePath"] != null ? (string)rowHead["LogFilePath"].ToString() : string.Empty;
                result.MailBody = rowHead["MailBody"] != null ? (string)rowHead["MailBody"].ToString() : string.Empty;
                result.MailCc = rowHead["MailCc"] != null ? (string)rowHead["MailCc"].ToString() : string.Empty;
                result.MailFrom = rowHead["MailFrom"] != null ? (string)rowHead["MailFrom"].ToString() : string.Empty;
                result.MailSmtp = rowHead["MailSmtp"] != null ? (string)rowHead["MailSmtp"].ToString() : string.Empty;
                result.MailStatus = rowHead["MailStatus"] != null ? (string)rowHead["MailStatus"].ToString() : string.Empty;
                result.MailSubject = rowHead["MailSubject"] != null ? (string)rowHead["MailSubject"].ToString() : string.Empty;
                result.MailTo = rowHead["MailTo"] != null ? (string)rowHead["MailTo"].ToString() : string.Empty;
                result.UplaodDirectory = rowHead["UplaodDirectory"] != null ? (string)rowHead["UplaodDirectory"].ToString() : string.Empty;
                result.UplaodFilePath = rowHead["UplaodFilePath"] != null ? (string)rowHead["UplaodFilePath"].ToString() : string.Empty;
                result.UplaodHost = rowHead["UplaodHost"] != null ? (string)rowHead["UplaodHost"].ToString() : string.Empty;
                result.UplaodPassword = rowHead["UplaodPassword"] != null ? (string)rowHead["UplaodPassword"].ToString() : string.Empty;
                result.UplaodPort = rowHead["UplaodPort"] != null ? (string)rowHead["UplaodPort"].ToString() : string.Empty;
                result.UplaodStatus = rowHead["UplaodStatus"] != null ? (string)rowHead["UplaodStatus"].ToString() : string.Empty;
                result.UplaodUsername = rowHead["UplaodUsername"] != null ? (string)rowHead["UplaodUsername"].ToString() : string.Empty;
            }

            return result;
        }
    }
}
