# Portfolio analityczne - Dawid Wolanin: automatyzacja VBA/Excel

Niniejszy projekt ma na celu zaprezentowanie automatyzacji przykładowych zadań w codziennej pracy w środowisku MS Excel przy wykorzystaniu VBA oraz korzystania z narzędzia UserForm.

# Spis treści

- [Importowanie z pliku txt](#importowanie-z-pliku-txt)
- [Automatyzacja raportu](#automatyzacja-raportu)
- [Formularz użytkownika UserForm](#userform)


```vba
Public Sub ImportTextFile()
    ' Ta podprocedura importuje jeden lub więcej plików .txt do nowych arkuszy w bieżącym skoroszycie Excel.
    
    ' Deklarowanie zmiennych
    Dim TextFile As Workbook         ' Zmienna do przechowywania obiektu skoroszytu dla otwieranego pliku tekstowego
    Dim OpenFiles() As Variant       ' Tablica do przechowywania ścieżek wybranych plików tekstowych
    Dim i As Integer                 ' Licznik pętli do iterowania po wybranych plikach
    
    ' Wywołaj funkcję GetFiles, aby wyświetlić okno dialogowe wyboru pliku i zachować ścieżki wybranych plików
    OpenFiles = GetFiles()
    
    ' Wyłącz aktualizację ekranu, aby poprawić wydajność i uniknąć migotania podczas procesu
    Application.ScreenUpdating = False
    
    ' Pętla po każdym wybranym pliku, importowanie jego zawartości do nowego arkusza
    For i = 1 To Application.CountA(OpenFiles)
        
        ' Otwórz plik tekstowy jako nowy skoroszyt i przypisz do niego zmienną TextFile
        Set TextFile = Workbooks.Open(OpenFiles(i))
        
        ' Kopiuj cały region danych z pierwszego arkusza pliku tekstowego (zakładając, że wszystko znajduje się w jednym regionie)
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        
        ' Aktywuj pierwotny skoroszyt (ten, który był otwarty przed importem plików tekstowych)
        Workbooks(1).Activate
        
        ' Dodaj nowy arkusz do oryginalnego skoroszytu
        Workbooks(1).Worksheets.Add
        
        ' Wklej skopiowane dane do nowo dodanego arkusza
        ActiveSheet.Paste
        
        ' Przenieś nowy arkusz do właściwej pozycji w skoroszycie
        ' Jest umieszczany za arkuszem o indeksie (i + 1), utrzymując kolejność importowanych plików
        
        ' Zmień nazwę nowego arkusza tak, aby odpowiadała nazwie oryginalnego pliku tekstowego, usuwając rozszerzenie ".txt"
        ActiveSheet.Name = Replace(TextFile.Name, ".txt", "")
        
        ' Wyjdź z trybu wycinania/kopiowania, aby wyczyścić schowek i uniknąć problemów z kolejnymi operacjami
        Application.CutCopyMode = False
        
        ' Zamknij skoroszyt pliku tekstowego bez zapisywania zmian (został otwarty tylko do kopiowania danych)
        TextFile.Close
    Next i
    
    ' Włącz ponownie aktualizację ekranu po zakończeniu procesu importu
    Application.ScreenUpdating = True
End Sub

Public Function GetFiles() As Variant
    ' Ta funkcja wyświetla okno dialogowe, które umożliwia użytkownikowi wybranie jednego lub więcej plików .txt do importu.
    ' Zwraca tablicę ścieżek wybranych plików.
    
    GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
End Function
```
