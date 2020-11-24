# Excel Chart!
# ODCINEK - 3 
# W tym odcinku pokażę jak możemy narysować dowolny typ wykresu w arkuszu Excela przy użyciu OpenXML.


# Krok - 1  
# Utworzymy teraz arkusz kalkulacyjny z danymi i nie użyjemy tutaj kodu z poprzednich części. Rozpoczynamy nowy projekt !!!
# Poniższy kod to kompletny kod do tworzenia arkusza Excela z danymi co już umiemy robić. 
# Powinno to być łatwe do zrozumienia. Jeśli nie, zapoznaj się z poprzednimi częściami tej serii.

# Kod ten inicjuje niektóre przykładowe dane przy użyciu klasy Person i używa tych danych do utworzenia arkusza kalkulacyjnego. 
# Do arkusza kalkulacyjnego nie zastosowano żadnych stylów żeby niepotrzebnie nie komplikować!


# Krok - 2
# I teraz W metodzie Main klasy Program utwórzymy obiekt naszej klasy ExcelChartTest i wywołamy metodę CreateExcelDocument, 
# przekazując ścieżkę do pliku.

# Możemy uruchomić program i zobaczyć plik excel z tabelką danych dla której stworzymy wykres!

##################################################################################################################################


# Krok - 3
# Narysujemy wykres w arkuszu kalkulacyjnym
# Po utworzeniu arkusza dodajmy DrawingsPart do arkusza i inicjujemy rysunek w arkuszu.

https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.drawingspart?redirectedfrom=MSDN&view=openxml-2.8.1

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


# Krok - 4
# Dodajemy wykres do DrawingPart.

		        // step 4
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });

                Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
                chart.AppendChild(new AutoTitleDeleted() { Val = true }); // We don't want to show the chart title
