using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoTestSelenium
{
    public class Log
    {
        private int index_row ;
        

        public string log_path;

        public Log()
        {
            string systemPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();
     

            if (!Directory.Exists(systemPath + "\\LogOutput"))
            {
                Directory.CreateDirectory(systemPath + "\\LogOutput");
            }

            log_path = systemPath + "\\LogOutput\\SeleniumCloudServices_Log_"+ DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"; 

        }

        public bool InitLog()
        {
            try
            {
                // Row starting
                index_row = 2;

                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(log_path, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Log"
                };
                sheets.Append(sheet);

                Row headerRow = new Row();

         
                Cell cell_date = new Cell() { CellReference = "B1", CellValue = new CellValue("Date"), DataType = CellValues.String };
                Cell cell_time = new Cell() { CellReference = "C1", CellValue = new CellValue("Time"), DataType = CellValues.String };
                Cell cell_operation = new Cell() { CellReference = "D1", CellValue = new CellValue("Operation"), DataType = CellValues.String };
                Cell cell_status = new Cell() { CellReference = "E1", CellValue = new CellValue("Status"), DataType = CellValues.String };

                headerRow.Append(cell_date);
                headerRow.Append(cell_time);
                headerRow.Append(cell_operation);
                headerRow.Append(cell_status);

                sheetData.Append(headerRow);
                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();

                return true;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public bool WriteLog(string operation, string status)
        {
            try
            {
                SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(log_path, true);
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, "Log");
          
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();


                Row row = new Row();
             
                Cell cell_date = new Cell() { CellReference = "B" +index_row, CellValue = new CellValue(DateTime.Now.ToString("yyyy/MM/dd")), DataType = CellValues.String };
                Cell cell_time = new Cell() { CellReference = "C"+index_row, CellValue = new CellValue(DateTime.Now.ToString("HH:mm:ss")), DataType = CellValues.String };
                Cell cell_operation = new Cell() { CellReference = "D"+index_row, CellValue = new CellValue(operation), DataType = CellValues.String };
                Cell cell_status = new Cell() { CellReference = "E"+index_row, CellValue = new CellValue(status), DataType = CellValues.String };

                row.Append(cell_date);
                row.Append(cell_time);
                row.Append(cell_operation);
                row.Append(cell_status);

                sheetData.Append(row);
                worksheet.Save();

                // Close the document.  
                spreadSheet.Close();
                index_row++;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                            Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                return null;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }


    }
}
