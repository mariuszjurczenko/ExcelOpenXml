using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
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

                // step 4
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });

                Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
                chart.AppendChild(new AutoTitleDeleted() { Val = true }); // We don't want to show the chart title

                // step 5
                PlotArea plotArea = chart.AppendChild(new PlotArea());
                Layout layout = plotArea.AppendChild(new Layout());

                BarChart barChart = plotArea.AppendChild(new BarChart(
                        new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                        new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) },
                        new VaryColors() { Val = false }
                ));

                // Constructing header
                Row row = new Row();
                int rowIndex = 1;
                // first empty
                row.AppendChild(ConstructCell(string.Empty, CellValues.String));
                foreach (var month in Months.Short)
                {
                    row.AppendChild(ConstructCell(month, CellValues.String));
                }
                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);
                rowIndex++;

                // step 6 
                for (int i = 0; i < people.Count; i++)
                {
                    BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
                        new Index() { Val = (uint)i },
                        new Order() { Val = (uint)i },
                        new SeriesText(new NumericValue() { Text = people[i].Name })
                    ));

                    // Adding category axis to the chart
                    CategoryAxisData categoryAxisData = barChartSeries.AppendChild(new CategoryAxisData());

                    // Category
                    // Constructing the chart category
                    string formulaCat = "People!$B$1:$M$1";

                    StringReference stringReference = categoryAxisData.AppendChild(new StringReference()
                    {
                        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
                    });

                    StringCache stringCache = stringReference.AppendChild(new StringCache());
                    stringCache.Append(new PointCount() { Val = (uint)Months.Short.Length });

                    for (int j = 0; j < Months.Short.Length; j++)
                    {
                        stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(Months.Short[j]));
                    }
                }

                var chartSeries = barChart.Elements<BarChartSeries>().GetEnumerator();


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
