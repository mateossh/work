Attribute VB_Name = "NewMacros"
Dim UndoHistory As UndoRecord

Function RemoveEmptyRows(cellIndex, rowIndex, table)
    If rowIndex <= table.Range.Rows.Count Then
        If table.Range.Rows.Item(rowIndex).Cells(cellIndex).Range.Text <> Chr(13) & Chr(7) Then
            'MsgBox ("W komórce są dane, skok do następnego wiersza")
            RemoveEmptyRows = RemoveEmptyRows(1, rowIndex + 1, table)
        
        ElseIf table.Range.Rows.Item(rowIndex).Cells(cellIndex).Range.Text = Chr(13) & Chr(7) Then
            'MsgBox ("BRAK DANYCH W KOMÓRCE !!!")
            
            If cellIndex < table.Range.Rows.Item(rowIndex).Cells.Count Then
                ' obecna komórka nie jest ostatnia
                
                ' skok do następnej
                RemoveEmptyRows = RemoveEmptyRows(cellIndex + 1, rowIndex, table)
            End If
            
            If cellIndex = table.Range.Rows.Item(rowIndex).Cells.Count Then
                'MsgBox ("ostatni komórka w wierszu, cały wiersz jest pusty USUWANIE !!!!!")
                table.Range.Rows.Item(rowIndex).Delete
                    
                ' sprawdzanie kolejnego wiersza
                RemoveEmptyRows = RemoveEmptyRows(1, rowIndex + 1, table)
            End If
        End If
    End If
    
End Function

Sub debestiofikator()
    On Error Resume Next ' Error 5991 workaround XD
    
    Set UndoHistory = Application.UndoRecord

    ' --- Ustawienie marginesów strony --------------
    UndoHistory.StartCustomRecord ("Ustawienie marginesów strony")
    
    With ActiveDocument.PageSetup
        .LeftMargin = CentimetersToPoints(1.8)
        .RightMargin = CentimetersToPoints(1.8)
        .BottomMargin = CentimetersToPoints(1.8)
        .TopMargin = CentimetersToPoints(2.5)
    End With
    
    UndoHistory.EndCustomRecord
    ' -----------------------------------------------
    
    
    ' --- Liczenie szerokości tabeli ----------------
    Dim TableWidth As Single
    If ActiveDocument.PageSetup.Orientation = wdOrientLandscape Then
        TableWidth = (297 - (18 * 2)) / 10
    ElseIf ActiveDocument.PageSetup.Orientation = wdOrientPortrait Then
        TableWidth = (210 - (18 * 2)) / 10
    End If
    ' -----------------------------------------------
    
    
    ' --- Style tabeli ------------------------------
    UndoHistory.StartCustomRecord ("Tworzenie stylu tabeli")
    
    Dim TableStyle As Style
    Set TableStyle = ActiveDocument.Styles.Add( _
        Name:="TableStyle 1", Type:=wdStyleTypeTable)
    
    With TableStyle
        .table.Alignment = wdAlignRowCenter
        .table.Borders.Enable = True
    End With
    
    UndoHistory.EndCustomRecord
    ' -----------------------------------------------
    
    
    ' --- Formatowanie tabeli i komórek -------------
    Dim t As table
    For Each t In ActiveDocument.Tables
        ' --------------------------------------------
    
        UndoHistory.StartCustomRecord ("Formatowanie tabeli i komórek")
        
        t.PreferredWidthType = wdPreferredWidthPoints
        t.PreferredWidth = CentimetersToPoints(TableWidth)
        t.Style = TableStyle
        t.Rows.HeightRule = wdRowHeightAtLeast
        t.Rows.Height = CentimetersToPoints(0.3)
        t.Range.Cells.VerticalAlignment = wdAlignVerticalCenter
        
        'Call RemoveEmptyRows(1, 1, t)
        
        For Each cell In t.Range.Cells
            With cell.Range
                .Font.Size = 5
                .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                .ParagraphFormat.LineSpacing = LinesToPoints(0.8)
            End With
        Next
        
        UndoHistory.EndCustomRecord
    Next
    ' -----------------------------------------------
End Sub
