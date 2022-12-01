Attribute VB_Name = "NewMacros"
Sub makro()
'
' makro Makro
'
'
    ' selection cells
    ' https://learn.microsoft.com/en-us/office/vba/api/word.selection.cells
    
    ' spliting strings by delimiter
    ' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/split-function
            
    ' dodawanie tekstu do komórki
    'Selection.Cells(7).Range.InsertBefore()
    
    'Application.Templates.LoadBuildingBlocks
    
    If Selection.Information(wdWithInTable) = True Then
        Dim ilosc_osob As String
        'Dim wysokosc_oplaty As String
        Dim wysokosc_zwolnienia As String
        Dim wynik As String
        
        ilosc_osob = Selection.Cells(1).Range.Text
        'wysokosc_oplaty = Selection.Cells(3).Range.Text
        wysokosc_zwolnienia = Selection.Cells(5).Range.Text
        wynik = Selection.Cells(7).Range.Text
        
        ' pierwsza komórka mia³a dziwny znak na koñcu
        Dim delimiter As String
        delimiter = Chr(13) & Chr(7)
        
        ' castowanie liczby osob do liczby
        Dim ilosc_osob_split() As String
        ilosc_osob_split = Split(ilosc_osob, delimiter)
        ilosc_osob_number = CInt(ilosc_osob_split(0))
        
        ' castowanie wysokosci oplaty do liczby (CCur)
        'Dim wysokosc_oplaty_split() As String
        'wysokosc_oplaty_split = Split(wysokosc_oplaty, " ")
        'wysokosc_oplaty_number = CCur(wysokosc_oplaty_split(0))
        
        ' castowanie wysokosci zwolnienia do liczby (CCur)
        'Dim wysokosc_zwolnienia_split() As String
        'wysokosc_zwolnienia_split = Split(wysokosc_zwolnienia, " ")
        'wysokosc_zwolnienia_number = CCur(wysokosc_zwolnienia_split(0))
        
        ' Val chce . zamiast , w stringu
        'Dim ilosc_osob, wysokosc_oplaty, wysokosc_zwolnienia As Currency
        'ilosc_osob = Val(Selection.Cells(1).Range.Text)
        'wysokosc_oplaty = Val(Selection.Cells(3).Range.Text)
        'wysokosc_zwolnienia = Val(Selection.Cells(5).Range.Text)
        
        
        Dim kwota_zwolnienia, wysokosc_oplaty, miesieczna_oplata As Currency
        
        kwota_zwolnienia = 1.75 * ilosc_osob_number
        wysokosc_oplaty = 17.5 * ilosc_osob_number
        miesieczna_oplata = wysokosc_oplaty - kwota_zwolnienia
        
        Selection.Rows(1).Cells(2).Range.Text = "x"
        Selection.Rows(1).Cells(3).Range.Text = wysokosc_oplaty & " z³"
        Selection.Rows(1).Cells(4).Range.Text = "-"
        Selection.Rows(1).Cells(5).Range.Text = kwota_zwolnienia & " z³"
        Selection.Rows(1).Cells(6).Range.Text = " "
        Selection.Rows(1).Cells(7).Range.Text = miesieczna_oplata & " z³"
        
        'MsgBox (wysokosc_oplaty)
        'MsgBox (miesieczna_oplata)
        
        ' U¿ywaj Selection.Rows(1).Cells(....). Bez Rows(1), indeksy komórek siê zmieniaj¹ (??)
        'MsgBox (Selection.Rows(1).Cells.Count)
        
        ' Przydatne do debugowania indeksów komórek
        'Selection.Cells(1).Shading.BackgroundPatternColorIndex = wdRed
    Else
        MsgBox "The insertion point is not in a table."
    End If



End Sub
