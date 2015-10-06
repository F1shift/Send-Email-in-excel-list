using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace SendEMail
{
    public class MailSender
    {
        public static bool SendMails(string excelFilePath, SmtpClient Smtp)
        {
            _Application excel;
            //try
            //{
                excel = OpenExcelSheet(excelFilePath);
                _Worksheet sheet = excel.Worksheets[1];
                int i = 2;
                String ErrorLog = "";
                MailMessage aMail = Read1MailFromExcel(sheet, i, (e) => { ErrorLog += e + "\r\n"; }); i++;
                while (aMail != null)
                {
                    try
                    {
                        Smtp.Send(aMail);
                    }
                    catch
                    {
                        ErrorLog += (i-1) + " Sending Error!\r\n";
                    }
                    aMail = Read1MailFromExcel(sheet, i, (e) => { ErrorLog += e + "\r\n"; }); i++;
                }
                if (!String.IsNullOrEmpty(ErrorLog))
                {
                    if (File.Exists("ErrorLog.txt"))
	                    File.Delete("ErrorLog.txt");
                    StreamWriter sw = new StreamWriter("ErrorLog.txt");
                    sw.Write(ErrorLog);
                    sw.Flush();
                    sw.Close();
                    Process nodepad = new Process();
                    nodepad.StartInfo.FileName = "ErrorLog.txt";
                    nodepad.Start();
                }
                if (excel != null)
                    excel.Quit();
                return true;
            //}
            //catch
            //{
            //    if (excel != null)
            //        excel.Quit();
            //    return false;
            //}
        }

        public static _Application OpenExcelSheet(string filePath)
        {
            //引用Excel Application類別
            _Application myExcel = null;
            //引用活頁簿類別 
            _Workbook myBook = null;

            //開啟一個新的應用程式
            myExcel = new Microsoft.Office.Interop.Excel.Application();

            //讓Excel文件可見 
            myExcel.Visible = false;

            myBook = myExcel.Workbooks.Open(filePath, ReadOnly: false);

            return myExcel;
        }

        public static MailMessage Read1MailFromExcel(_Worksheet excelSheet, int rowIndex, Action<String> ErrorOutput)
        {
            if (excelSheet.Cells[rowIndex, 1].Value2 != null)
            {
                MailMessage newMail = new MailMessage();
                //加入寄件者
                newMail.From = new MailAddress(((String)excelSheet.Cells[rowIndex, 1].Value2).Trim());
                //加入收件者
                if (excelSheet.Cells[rowIndex, 2].Value2 != null)
                {
                    String[] Tos = ((String)excelSheet.Cells[rowIndex, 2].Value2).Split(',', ';', '\n');
                    foreach (var to in Tos)
                    {
                        MailAddress adress = null;
                        try
                        {
                            adress = new MailAddress(to.Trim());
                            newMail.To.Add(adress);
                        }
                        catch
                        {
                            if (ErrorOutput != null)
                                ErrorOutput(rowIndex + " Something wrong with Mail Adress : " + to);
                        }
                    }
                }
                //加入副本
                if (excelSheet.Cells[rowIndex, 3].Value2 != null)
                {
                    String[] CCs = ((String)excelSheet.Cells[rowIndex, 3].Value2).Split(',', ';', '\n');
                    foreach (var to in CCs)
                    {
                        MailAddress adress = null;
                        try
                        { 
                            adress = new MailAddress(to.Trim());
                            newMail.CC.Add(adress);
                        }
                        catch
                        {
                            if (ErrorOutput != null)
                                ErrorOutput(rowIndex + " Something wrong with Mail Adress : " + to);
                        }
                    }
                }
                //加入密件副本
                if (excelSheet.Cells[rowIndex, 4].Value2 != null)
                {
                    String[] BCCs = ((String)excelSheet.Cells[rowIndex, 4].Value2).Split(',', ';', '\n');
                    foreach (var to in BCCs)
                    {
                        MailAddress adress = null;
                        try
                        {
                            adress = new MailAddress(to.Trim());
                            newMail.Bcc.Add(adress);
                        }
                        catch
                        {
                            if (ErrorOutput != null)
                                ErrorOutput(rowIndex + " Something wrong with Mail Adress : " + to);
                        }
                    }
                }
                //信件標題
                if (excelSheet.Cells[rowIndex, 5].Value2 != null)
                {
                    String Subject = excelSheet.Cells[rowIndex, 5].Value2;
                    newMail.Subject = Subject;
                }
                else
                    if (ErrorOutput != null)
                        ErrorOutput(rowIndex + " Subject is null!");
                //信件內容
                if (excelSheet.Cells[rowIndex, 6].Value2 != null)
                {
                    String Body = excelSheet.Cells[rowIndex, 6].Value2;
                    newMail.Body = Body;
                }
                else
                    if (ErrorOutput != null)
                        ErrorOutput(rowIndex + " Body is null!");
                //附件
                String[] Attachments = ((String)excelSheet.Cells[rowIndex, 7].Value2).Split(',', ';', '\n');
                if (excelSheet.Cells[rowIndex, 7].Value2 != null)
                {
                    //[] Attachments = ((String)excelSheet.Cells[rowIndex, 7].Value2).Split(',', ';', '\n');
                    Attachments = (from filepath in Attachments
                                  select filepath.Trim()).ToArray();
                    foreach (var fileName in Attachments)
                        if (File.Exists(@fileName))
                        {
                            newMail.Attachments.Add(new Attachment(File.Open(@fileName, FileMode.Open, FileAccess.Read, FileShare.Read), Path.GetFileName(@fileName)));
                        }
                        else if (File.Exists(fileName))
                        {
                            newMail.Attachments.Add(new Attachment(File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.Read), Path.GetFileName(fileName)));
                        }
                        else
                        {
                            if (ErrorOutput != null)
                                ErrorOutput(rowIndex + " File not found : " + fileName + "!");
                        }
                }
                return newMail;
            }
            else
            {
                return null;
            }
        }
    }
}
