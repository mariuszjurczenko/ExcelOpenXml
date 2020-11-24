using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace ExcelChart
{
    class ExcelChartTest
    {
        public void CreateExcelDoc(string fileName)
        {
            List<Person> people = new List<Person>();
            Initizalize(people);

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "People" };
                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // step 3
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();




                sheets.Append(sheet);
                workbookPart.Workbook.Save();

              
                // Constructing header
                Row row = new Row();
                // first empty
                row.AppendChild(ConstructCell(string.Empty, CellValues.String));
                foreach (var month in Months.Short)
                {
                    row.AppendChild(ConstructCell(month, CellValues.String));
                }
                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting people
                foreach (var person in people)
                {
                    row = new Row();
                    row.AppendChild(ConstructCell(person.Name, CellValues.String));

                    foreach (var value in person.Values)
                    {
                        row.AppendChild(ConstructCell(value.ToString(), CellValues.Number));
                    }
                    sheetData.AppendChild(row);
                }
                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
            };
        }

        private void Initizalize(List<Person> people)
        {
            people.AddRange(new Person[] {
                new Person
                {
                    Name = "Marcin",
                    Values = new byte[] { 14, 25, 29, 18, 21, 17, 26, 24, 19, 21, 28, 24 }
                },
                new Person
                {
                    Name = "Mariusz",
                    Values = new byte[] { 20, 15, 26, 18, 21, 17, 26, 24, 19, 30, 10, 15 }
                },
                new Person
                {
                    Name = "Tomek",
                    Values = new byte[] {  18, 22, 24, 18, 30, 10, 19, 22, 15, 27, 18, 23 }
                }
            });
        }
    }
}
