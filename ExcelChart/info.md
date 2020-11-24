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



# Krok - 5
# Dodajemy PlotArea do wykresu i dołącz Layout oraz BarChart jako jego elementy podrzędne.

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
        -->     rowIndex++;



# Krok - 6
# Dodajemy serię i kategorie dla wykresu
# Po skonstruowaniu wiersza nagłówka, dla każdej osoby dodamy ChartSeries do BarChart.
# Dla każdego BarSeries dodajemy komórki odniesienia w arkuszu kalkulacyjnym, 
# tworząc formułę Studenci! $ B $ 0: $ G $ 0. 
# Po dodaniu referencji utworzymy StringCache dla rzeczywistych danych.

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



# Krok - 7
# Dodajemy wartości dla wykresu
# Dla każdej osoby dodajemy rzeczywiste wartości do każdej serii. 
# Zwróć uwagę, że tak samo jak w przypadku kategorii dodajemy odniesienie do danych w arkuszu kalkulacyjnym za pomocą formuły, 
# a także dodajemy rzeczywiste dane do pamięci podręcznej.          
# !!!!  usuwamy // Inserting people

                // step 7
                for (int i = 0; i < people.Count; i++)
                {
                    row = new Row();

                    row.AppendChild(ConstructCell(people[i].Name, CellValues.String));

                    chartSeries.MoveNext();

                    string formulaVal = string.Format("People!$B${0}:$M${0}", rowIndex);
                    DocumentFormat.OpenXml.Drawing.Charts.Values values = chartSeries.Current.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

                    NumberReference numberReference = values.AppendChild(new NumberReference()
                    {
                        Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
                    });

                    NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
                    numberingCache.Append(new PointCount() { Val = (uint)Months.Short.Length });

                    for (uint j = 0; j < people[i].Values.Length; j++)
                    {
                        var value = people[i].Values[j];

                        row.AppendChild(ConstructCell(value.ToString(), CellValues.Number));

                        numberingCache.AppendChild(new NumericPoint() { Index = j }).Append(new NumericValue(value.ToString()));
                    }

                    sheetData.AppendChild(row);
                    rowIndex++;
                }

                barChart.AppendChild(new DataLabels(
                                    new ShowLegendKey() { Val = false },
                                    new ShowValue() { Val = false },
                                    new ShowCategoryName() { Val = false },
                                    new ShowSeriesName() { Val = false },
                                    new ShowPercent() { Val = false },
                                    new ShowBubbleSize() { Val = false }
                                ));

                barChart.Append(new AxisId() { Val = 48650112u });
                barChart.Append(new AxisId() { Val = 48672768u });



# Krok - 8
# Konfigurowanie CategoyAxis i ValueAxis
# Do tej pory dodaliśmy rysunek do arkusza i dołączaliśmy do niego wykres. 
# Następnie dodaliśmy ChartSeries do Chart. Rzeczywista kategoria i wartości zostały dodane do każdej ChartSeries. 
# Teraz musimy ustawić osie kategorii i wartości!!!

                // step 8
                // Adding Category Axis
                plotArea.AppendChild(
                    new CategoryAxis(
                        new AxisId() { Val = 48650112u },
                        new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                        new Delete() { Val = false },
                        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                        new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                        new CrossingAxis() { Val = 48672768u },
                        new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                        new AutoLabeled() { Val = true },
                        new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
                    )
                );

                // Adding Value Axis
                plotArea.AppendChild(
                    new ValueAxis(
                        new AxisId() { Val = 48672768u },
                        new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                        new Delete() { Val = false },
                        new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                        new MajorGridlines(),
                        new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                        {
                            FormatCode = "General",
                            SourceLinked = true
                        },
                        new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                        new CrossingAxis() { Val = 48650112u },
                        new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                        new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
                    )
                );

                chart.Append(
                    new PlotVisibleOnly() { Val = true },
                    new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                    new ShowDataLabelsOverMaximum() { Val = false }
                );

                chartPart.ChartSpace.Save();  




# Krok 9
# TwoCellAnchor
# Na koniec musimy zdefiniować, gdzie fizycznie ma się pojawić wykres. Na tym polega odpowiedzialność TwoCellAnchor.

                // step 9
                // Positioning the chart on the spreadsheet
                TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                        new ColumnId("0"),
                        new ColumnOffset("0"),
                        new RowId((rowIndex + 2).ToString()),
                        new RowOffset("0")
                ));

                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                        new ColumnId("8"),
                        new ColumnOffset("0"),
                        new RowId((rowIndex + 12).ToString()),
                        new RowOffset("0")
                ));

                // Append GraphicFrame to TwoCellAnchor
                GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new GraphicFrame());
                graphicFrame.Macro = string.Empty;

                graphicFrame.Append(new NonVisualGraphicFrameProperties(
                        new NonVisualDrawingProperties()
                        {
                            Id = 2u,
                            Name = "Sample Chart"
                        },
                        new NonVisualGraphicFrameDrawingProperties()
                ));

                graphicFrame.Append(new Transform(
                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
                ));

                graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                            )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                 ));

                twoCellAnchor.Append(new ClientData());

                

# i na koniec zapisujem rysunek  (Drawing Part).

                drawingsPart.WorksheetDrawing.Save();
                worksheetPart.Worksheet.Save();