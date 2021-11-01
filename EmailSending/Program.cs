using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using LumenWorks.Framework.IO.Csv;
using OfficeOpenXml;
using Renci.SshNet;
using Spire.Xls;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailSending
{
    class Program
    {
        static void Main(string[] args)
        {
            //try
            //{
            //    MailMessage message = new MailMessage();
            //    SmtpClient smtp = new SmtpClient();
            //    message.From = new MailAddress("testuser4312@gmail.com");
            //    message.To.Add(new MailAddress("jaydipmer12345@gmail.com"));
            //    message.Subject = "Test";
            //    message.IsBodyHtml = true; //to make message body as html  
            //    message.Body = "Test";
            //    smtp.Port = 587;
            //    smtp.Host = "smtp.gmail.com"; //for gmail host  
            //    smtp.EnableSsl = true;
            //    smtp.UseDefaultCredentials = false;
            //    smtp.Credentials = new NetworkCredential("testuser4312@gmail.com", "Testuser@4312");
            //    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    smtp.Send(message);
            //}
            //catch (Exception e) { }



            //below code is for get file from the sftp server

            try
            {



                String Host = "devsftp.anewbiz.net";
                int Port = 2222;
                String RemoteFileName = @"Deploy/main.2021.09.27.csv";
                String LocalDestinationFilename = @"D:\main.2021.09.27.csv";
                String Username = "devteam";
                String Password = "@newb1z!";

                using (var sftp = new SftpClient(Host, Port, Username, Password))
                {
                    sftp.Connect();

                    using (var file = File.OpenWrite(LocalDestinationFilename))
                    {
                        sftp.DownloadFile(RemoteFileName, file);
                    }

                    sftp.Disconnect();
                }
            }
            catch (Exception e)
            {

            }





            // below code is for the conversation of csv to excel

            try
            {

                for (int i = 0; i < 2; i++)
                {
                    string fileName = i == 0 ? @"C:\Users\jaydip mer\HomeDemotFielsConverstaion\main.csv" : @"C:\Users\jaydip mer\HomeDemotFielsConverstaion\details.csv";
                    string sheetName = i == 0 ? "main" : "details";
                    string ResultFile = i == 0 ? @"C:\Users\jaydip mer\HomeDemotFielsConverstaion\Main.xlsx" :
                        @"C:\Users\jaydip mer\HomeDemotFielsConverstaion\details.xlsx";
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(fileName, ",");
                    workbook.LoadFromFile(@"C:\Users\jaydip mer\Excel_Files\Inventory.xls");
                    Worksheet sheet = workbook.Worksheets[0];

                    sheet.Name = sheetName;


                    workbook.SaveToFile(ResultFile, ExcelVersion.Version2010);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    var workbookFileInfo = new FileInfo(ResultFile);
                    using (var excelPackage = new ExcelPackage(workbookFileInfo))
                    {
                        var worksheet = excelPackage.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Evaluation Warning");
                        excelPackage.Workbook.Worksheets.Delete(worksheet);
                        excelPackage.Save();
                    }
                }



            }
            catch (Exception e)
            {

            }





            //below code is for merge files.
            try
            {
                Workbook newbook = new Workbook();
                newbook.Version = ExcelVersion.Version2010;
                newbook.Worksheets.Clear();

                Workbook tempbook = new Workbook();
                string[] excelFiles = new String[] { "main.xlsx", "details.xlsx" };
                for (int i = 0; i < excelFiles.Length; i++)
                {
                    string path = @"C:\Users\jaydip mer\HomeDemotFielsConverstaion\" + excelFiles[i];
                    tempbook.LoadFromFile(path);
                    foreach (Worksheet sheet in tempbook.Worksheets)
                    {
                        newbook.Worksheets.AddCopy(sheet);
                    }
                }

                newbook.SaveToFile(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\Inventory.xlsx", ExcelVersion.Version2010);
                newbook.ActiveSheetIndex = 0;
                //System.Diagnostics.Process.Start(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\Inventory.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var workbookFileInfo = new FileInfo(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\Inventory.xlsx");
                using (var excelPackage = new ExcelPackage(workbookFileInfo))
                {
                    var worksheet = excelPackage.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Evaluation Warning");
                    excelPackage.Workbook.Worksheets.Delete(worksheet);
                    excelPackage.Save();
                }
            }
            catch (Exception e)
            {

            }

            // below code is for the Change the header of the excel

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\Inventory.xlsx");
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
                int columnCount = xlWorksheet.UsedRange.Columns.Count;
                List<string> columnNames = new List<string>();
                for (int c = 1; c < columnCount; c++)
                {
                    string name = xlWorksheet.Cells[1, c].Value;
                    switch (name)
                    {
                        case "cnt_whs_id":
                            xlWorksheet.Cells[1, c].Value = "WareHouse";
                            break;
                        case "GRP_NM":
                            xlWorksheet.Cells[1, c].Value = "Group ID";
                            break;
                        case "SBNUM":
                            xlWorksheet.Cells[1, c].Value = "SB Number";
                            break;
                        case "SBN_CREATED_ON":
                            xlWorksheet.Cells[1, c].Value = "SB Create Date";
                            break;
                        case "SALVAGE_CATEGORY":
                            xlWorksheet.Cells[1, c].Value = "Salvage Category";
                            break;
                        case "IN_AR":
                            xlWorksheet.Cells[1, c].Value = "In ARS";
                            break;
                        case "ORG_SALE_TYPE":
                            xlWorksheet.Cells[1, c].Value = "HdDotCom";
                            break;
                        case "STATUS":
                            xlWorksheet.Cells[1, c].Value = "Status";
                            break;
                        case "PALLET_ID":
                            xlWorksheet.Cells[1, c].Value = "Pallet ID";
                            break;
                        case "SBN_close_date":
                            xlWorksheet.Cells[1, c].Value = "SB Closed Date";
                            break;
                        default:

                            break;

                    }

                    //if(name== "cnt_whs_id")
                    //{
                    //    xlWorksheet.Cells[1, 1].Value = "WareHouse";
                    //    xlWorkbook.SaveAs(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\InventoryHeaderChange.xlsx");
                    //}



                }
                xlWorkbook.SaveAs(@"C:\Users\jaydip mer\HomeDemotFielsConverstaion\InventoryHeaderChange2.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }




        }
    }
}
