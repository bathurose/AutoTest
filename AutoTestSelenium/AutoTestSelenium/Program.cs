using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using System.IO;
using System.Windows;

using NUnit.Framework;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AutoTestSelenium
{
    [TestFixture]
    public class Program
    {
        [Test]
        public void Main()
        {
            string user, password, path;

            string systemPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();

            Log log = new Log();

            if (log.InitLog() == false)
            {
                Console.WriteLine("Cannot create the log file");
                return;
            }

            if (!(File.Exists(systemPath + "\\CloudServices_Information\\CloudServices.xlsx")))
            {
                log.WriteLog("ファイルエクセルクラウドサービス情報が存在しません", null);
                Console.WriteLine("ファイルエクセルクラウドサービス情報が存在しません");
             
                return;
            }

            // Excelファイル情報クラウドサービスへのパス
            path = systemPath + "\\CloudServices_Information\\CloudServices.xlsx";

            SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false);

            WorkbookPart workbookPart = doc.WorkbookPart;
   
            Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById("rId1")).Worksheet;
      
            SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();
             
            Cell theCellUser = (Cell)thesheetdata.ElementAt(1).ChildElements.ElementAt(4);
                          
            SharedStringItem textUser = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellUser.InnerText));
            user = textUser.Text.Text;

            Cell theCellPassword = (Cell)thesheetdata.ElementAt(1).ChildElements.ElementAt(5);

            SharedStringItem textPassword = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellPassword.InnerText));
            password = textPassword.Text.Text;

            // クラウドサービスのリストを取得する
            List<string> listCloudService = new List<string>();

            // クラウドサービスのステータスのリストを取得する 
            for (int i = 1; i < thesheetdata.ChildElements.Count()-2; i++)
            {
                Cell theCellCloudService = (Cell)thesheetdata.ElementAt(i).ChildElements.ElementAt(2);
                SharedStringItem textCloudService = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellCloudService.InnerText));
                listCloudService.Add(textCloudService.Text.Text);               
            }

            var stateStarting = "stop";

            MyTestCaseTest mst = new MyTestCaseTest();

            mst.SetUp();

            mst.StartAzure(user, password);

            for (int i = 0; i < listCloudService.Count; i++)
            {
                if (mst.TestCases(listCloudService[i], stateStarting) == true)
                {
                    log.WriteLog(stateStarting + " " + listCloudService[i], "OK");
                    Console.WriteLine(stateStarting + " " + listCloudService[i] + " " + "OK");
                }
                else
                {
                    log.WriteLog(stateStarting + " " + listCloudService[i], "ERROR");
                    Console.WriteLine(stateStarting + " " + listCloudService[i] + " " + "ERROR");
                }
            }
            mst.TearDown();
            

        }
    }
}

