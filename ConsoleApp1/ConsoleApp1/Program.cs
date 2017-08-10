using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;



namespace ConsoleApp1
{
    class Program
    {
        // The DOM approach.
        // Note that the code below works only for cells that contain numeric values.
        // The SAX approach.
        static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.WriteLine(reader.ElementType + "\t" + text);
                    }
                    else if (reader.ElementType == typeof(SharedStringTablePart.share))
                    {
                        text = reader.GetText();
                        Console.WriteLine(reader.ElementType + "\t" + text);
                    }
                }
            }
        }
        static void Main(string[] args)
        {
            Console.WriteLine("Hello word!");
            String fileName = @"C:\Users\Administrator\Desktop\test.xlsx";
            ReadExcelFileSAX(fileName);
            SpreadsheetDocument output =
                SpreadsheetDocument.Create(@"C:\Users\Administrator\Desktop\out.xlsx",
                    SpreadsheetDocumentType.Workbook);


            Console.ReadKey();
        }
    }
}
