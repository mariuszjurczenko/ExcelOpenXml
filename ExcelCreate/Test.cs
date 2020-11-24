using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelCreate
{
    public class Test
    {
        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = GenerateStyleshett();
                stylesPart.Stylesheet.Save();

                Columns columns = new Columns(
                    new Column { Min = 1, Max = 1, Width = 30, CustomWidth = true },
                    new Column { Min = 2, Max = 5, Width = 20, CustomWidth = true },
                    new Column { Min = 6, Max = 6, Width = 40, CustomWidth = true },
                    new Column { Min = 7, Max = 9, Width = 15, CustomWidth = true }
                 );

                worksheetPart.Worksheet.AppendChild(columns);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Arkusz 1" };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                WriteingToExcel(worksheetPart);

            }
        }

        private static void WriteingToExcel(WorksheetPart worksheetPart)
        {
            SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
            Row row = new Row();
            row.Append(new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 2 },
                       new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 2 },
                       new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 2 });

            sheetData.AppendChild(row);

            for (int i = 0; i < 20; i++)
            {
                row = new Row();
                row.Append(new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test " + i), DataType = CellValues.String, StyleIndex = 1 },
                           new Cell() { CellValue = new CellValue("napis test "), DataType = CellValues.String, StyleIndex = 1 });

                sheetData.AppendChild(row);
            }

            worksheetPart.Worksheet.Save();
        }

        private Stylesheet GenerateStyleshett()
        {
            Stylesheet stylesheet = null;

            Fonts fonts = new Fonts(
                new Font(
                    new FontSize() { Val = 10 }
                    ),
                new Font(
                    new FontSize() { Val = 16 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }
                    )
                );

            Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }),
                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
                new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                { PatternType = PatternValues.Solid })
            );

            Borders borders = new Borders(
                new Border(),   // index 0 default
                new Border(     // index 1 black border
                    new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Dotted },
                    new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Dotted },
                    new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            );

            CellFormats cellFormats = new CellFormats(
                new CellFormat(),
                new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true },
                new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }
            );

            stylesheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return stylesheet;
        }
    }
}
