# Pokaże teraz jak skonstruować plik Excel w C#
# W tym samouczku użyjemy biblioteki  Open XML, C# i .NET Core
# Zaczynajmy

# 1 Jestesmy w Visual Studio 2019 i utwórzmy nową aplikację konsolową C#

# 2 Teraz musimy dodać odniesienie do biblioteki:
    DocumentFormat.OpenXml

# 3 Teraz Dodajemy nową klasę i nazwijemy ja „Test” 
# i utwórzmy publiczną metodę o nazwie „CreateExcelDoc”. 
# W tej metodzie utworzymy i zapiszemy nasz plik Excel.

    public void CreateExcelDoc(string fileName)
    {  }

# 4 teraz Zaimportujemy następujące przestrzenie nazw do klasy:

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

# 5 Teraz Utworzymy nowy dokument arkusza kalkulacyjnego i przekaż nazwę pliku i dokument jako parametry.

    using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
    {  }

# 6 SpreadsheetDocument musi mieć co najmniej część WorkbookPart i WorkSheetPart.
#   Dodamy następujący kod w bloku using.

            // Dodajemy WorkbookPart do dokumentu.              Add a WorkbookPart to the document.
    WorkbookPart workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();
            // Dodajemy WorksheetPart do WorkbookPart.             Add a WorksheetPart to the WorkbookPart.
    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(new SheetData());

# 7 Wszystkie elementy arkusza kalkulacyjnego będą powiązane relacją rodzic / dziecko.
# Skoroszyt będzie zawierał nasze arkusze kalkulacyjnego.
# Arkusz będzie zawierał SheetData i Columns. Są to dane arkusza, 
# w których rzeczywiste wartości są umieszczane w wierszach i komórkach. 

# Inicjując arkusz roboczy, możemy dołączyć SheetData jako jego dziecko, przekazując go jako argument.
# Dołącz „Arkusze” do skoroszytu. Arkusze będą zawierać jeden lub wiele arkuszy, 
# z których każdy jest powiązany z częścią arkusza roboczego.

    Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

# 8 Następnie możemy dodać jeden lub wiele „Arkuszy” do Arkuszy. 
# To będą nasze arkusze excelsheets. Zwróć uwagę, że arkusz jest powiązany z WorksheetPart.

    Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Arkusz 1" };
    sheets.Append(sheet);

# 9 Na koniec zapisz skoroszytu.

    workbookPart.Workbook.Save();

# 10 I teraz W metodzie Main klasy Program utwórzymy obiekt naszej klasy test i wywołamy metodę CreateExcelDocument, 
# przekazując ścieżkę do pliku.

        static void Main(string[] args)
        {
            Test test = new Test();
            test.CreateExcelDoc(@"C:\Excel\test.xlsx");
        }

# 11 Uruchommy projekt i sprawdź wygenerowany plik Excel.

# 12 teraz jesli chcemy cos zapisac 
#    Dołącz SheetData do arkusza. SheetData działa jako kontener, do którego będą trafiać wszystkie wiersze i kolumny.

    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

# 13 Klasa Row reprezentuje wiersz w arkuszu kalkulacyjnym programu Excel. 
# Każdy wiersz może zawierać jedną lub więcej komórek. 
# Każda komórka będzie miała CellValue, która zawiera rzeczywistą wartość w komórce.

        SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
        Row row = new Row();
        Cell cell = new Cell() { CellValue = new CellValue("test"), DataType = CellValues.String };
        row.Append(cell);

        SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
        Row row = new Row();
        row.Append(new Cell() { CellValue = new CellValue("test"), DataType = CellValues.String });

# 14 Dodaj wiersz do arkusza kalkulacyjnego.

    sheetData.AppendChild(row);

#15 I  Zapisujemy arkusz na końcu.
    worksheetPart.Worksheet.Save();


#16
                for (int i = 1; i < 20; i++)
                {
                    row = new Row();
                    row.Append(new Cell() { CellValue = new CellValue("test " + i), DataType = CellValues.String },
                               new Cell() { CellValue = new CellValue("test rer " + i), DataType = CellValues.String });
                    sheetData.AppendChild(row);
                }




###########################################################################################################################

# W tym odcinku będziemy pracować z nasnaszym kodem z poprzedniego odcinka  i dodamy styl i dokonamy pewnych dostosowań naszego arkusza!!!


    - Krok 1

# Dodamy arkusz stylów do Exela

# Klasa Stylesheet służy do dodawania niestandardowego stylu do arkusza kalkulacyjnego. Arkusz stylów może akceptować różne elementy, 
# takie jak obramowania, kolory, wypełnienia itp. Jako elementy podrzędne, które określają wygląd arkusza kalkulacyjnego.

# CellFormats przechowuje kombinację różnych stylów, które później można zastosować w komórce.

# Aby dodać arkusz stylów do naszego skoroszytu, musimy dodać WorkbookStylePart do części skoroszytu i zainicjować jego właściwość StyleSheet.
# Zamierzamy uczynić nasz nagłówek pogrubionym i białym z ciemnym tłem, a także dodać obramowanie do wszystkich innych komórek.


# Utwórz nową metodę GenerateStylesheet (), która zwraca obiekt StyleSheet.

    private Stylesheet GenerateStylesheet()
    {
        Stylesheet styleSheet = null;   
        return styleSheet;
    }

--------------------------------------------------------------------

# Utworzymy teraz Czcionki - Fonty
# Czcionki mogą mieć jedno lub więcej elementów podrzędnych Font, z których każdy ma inne właściwości, takie jak FontSize, Bold, Color itp.
# Dodam teraz kod w metodzie GenerateStyleSheet:

Fonts fonts = new Fonts(
    new Font( // Index 0 - default
        new FontSize() { Val = 10 }

    ),
    new Font ( // Index 1 - header
        new FontSize() { Val = 14 },
        new Bold(),
        new Color() { Rgb = "FFFFFF" } //biały

    )
);
# Zwróć uwagę, że dodajemy dwa elementy podrzędne Font do obiektu Fonts. 
# Pierwsza z nich to domyślna czcionka używana przez wszystkie komórki, a druga jest specyficzna dla nagłówka.

--------------------------------------------------------------------

# teraz Wypełnienia
#Wypełnienia mogą mieć co najmniej jeden element podrzędny Fill, dla którego można ustawić jego kolor ForegroundColor.

Fills fills = new Fills(
         new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
         new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
         new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
         { PatternType = PatternValues.Solid }) // Index 2 - header
);
# Excel musi mieć domyślnie pierwsze dwa. Trzeci to styl, jaki chcemy mieć dla naszych komórek nagłówkowych; szare pogrubione tło.

--------------------------------------------------------------------

# Obramowania - Borders
# Obramowania mogą mieć jedno lub więcej elementów podrzędnych Border, z których każde określa, jak powinno wyglądać obramowanie:

Borders borders = new Borders(
        new Border(), // index 0 default
        new Border( // index 1 black border
            new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Dotted },
            new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Dotted },
            new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
            new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
            new DiagonalBorder())
);
# Pierwsza to ramka domyślna, a druga to nasza niestandardowa ramka.

--------------------------------------------------------------------

# CellFormats
# Teraz, gdy ustawilismy nasze niestandardowe formatowanie, możemy utworzyć CellFormats, które mają jedno lub wiele elementów potomnych CellFormat. 
# A Każdy CellFormat otrzymuje indeks Czcionki, Obramowania, Wypełnienie itp., Z którym będzie skojarzony:

CellFormats cellFormats = new CellFormats(
        new CellFormat(), // default
        new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
        new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
);


# Na koniec inicjalizujemy obiekt Stylesheet:

styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);


    - Krok 2

# Krok 2: Dodajemy styl do skoroszytu
# Dodajemy WorkbookStylePart do WorkbookPart i zainicjuj jego arkusz stylów:
                                                 worksheetPart.Worksheet = new Worksheet();
// Adding style
WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
stylePart.Stylesheet = GenerateStylesheet();
stylePart.Stylesheet.Save();


#Krok 3: Dodajemy styl do komórek

# Teraz, gdy mamy link do arkusza stylów do skoroszytu, możemy określić, jaki styl ma przestrzegać każda komórka.

# Każda Cell ma właściwość o nazwie StyleIndex, która pobiera indeks stylu, który chcemy zastosować do tej komórki. 
# Indeks tutaj odnosi się do indeksu CellFormats.
# Modyfikuje i przekazuje pożądany indeks stylu:

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


# Krok 4: możemy dodać niestandardową szerokość kolumn

# Dodanie niestandardowej szerokości do określonych kolumn jest bardzo łatwe. 
# Najpierw musimy utworzyć obiekt Columns, a następnie dodać co najmniej jedną kolumnę jako jego elementy potomne, 
# z których każdy będzie definiował niestandardową szerokość zakresu kolumn w arkuszu kalkulacyjnym. 
# Możesz zbadać inne właściwości kolumny, aby określić większe dostosowanie kolumn. Tutaj interesuje nas tylko określenie szerokości kolumn.

// Setting up columns
Columns columns = new Columns(
        new Column 
        {
            Min = 1,
            Max = 50,
            Width = 15,             // Szerokość kolumny
            CustomWidth = true      
        },
        new Column 
        {
            Min = 1,
            Max = 50,
            Width = 20,
            CustomWidth = true
        },
        new Column 
        {
            Min = 1,
            Max = 50,
            Width = 25,
            CustomWidth = true
        }
        ...
        );


# Na koniec musimy dołączyć kolumny do arkusza roboczego.

workheetPart.Worksheet.AppendChild(columns);



#Wynik
Uruchom aplikację i sprawdź wygenerowany plik Excel