''Module that contains useful functions used for all modules.

'Open the specified workbook if it isn´t already opened
'return the object corresponding to the workbook.
Function getWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook
    sFile = Dir(sFullName)
    On Error Resume Next
        Set wbReturn = Workbooks(sFile)

        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
        End If
    On Error GoTo 0

    Set getWorkbook = wbReturn
End Function

Function existBookInWB(wsheets As Sheets, name As String)
    On Error GoTo err
    existBookInWB = True
    f = IsObject(wsheets.Item(name))
    Exit Function
err:
    existBookInWB = False
End Function

'Function that return a datarange in string format by an index.
'1: Inar Total
'2: Detalle lineas adicionales
Function getRangeByIndex(ind As Integer) As String
    Dim e_wsheet As Worksheet
    Dim startpoint As Range
    Dim datarange As Range
    
    'Dim inar_tp_db_wb As Workbook
    'Set inar_tp_db_wb = getWorkbook(mainpath & "/Inar/Inar TP Junio2018.xlsx")

    Select Case ind
        Case 1
            'Set e_wsheet = inar_tp_db_wb.Worksheets("Inar TP Ordenado")
            Set e_wsheet = ActiveWorkbook.Worksheets("Inar Total")
            Set startpoint = e_wsheet.Range("A1")
            Set datarange = e_wsheet.Range("A1", startpoint.SpecialCells(xlLastCell))
        Case 2
            Set e_wsheet = ActiveWorkbook.Worksheets("Detalle lineas adicionales")
            Set startpoint = e_wsheet.Range("A1")
            Set datarange = e_wsheet.Range("A1", startpoint.SpecialCells(xlLastCell))
        Case Else
            MsgBox "Bad pivottable index"
    End Select

    getRangeByIndex = e_wsheet.name & "!" & datarange.Address(ReferenceStyle:=xlR1C1)

End Function

'Adds a new sheet after the last current sheet and set the name,
'returns the new sheet created.
Function addWorkSheetByName(wb As Workbook, name As String) As Worksheet
    Dim new_ws As Worksheet
    c = wb.Worksheets.Count
    wb.Worksheets.Add after:=wb.Worksheets(c)
    Set new_ws = wb.Worksheets(c + 1)
    
    If new_ws.name <> name Then
        new_ws.name = name
    End If
    'MsgBox TypeName(new_ws)
    Set addWorkSheetByName = new_ws
End Function


'Function that search by a header name in the first row of a worksheet.
'If a match happens then the corresponding cell is returned as a range object.
'Otherwise the 'Nothing' object is returned
Function searchHeaderByName(ws As Worksheet, name As String) As Range
    Dim header As Range
    finalcol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    matched = False

    For i = 1 To finalcol
        If name = ws.Cells(1, i) Then
            matched = True
            Set header = ws.Cells(1, i)
            Set searchHeaderByName = header
            Exit For
        End If
    Next i

    If matched = False Then
        MsgBox "Header " & name & " not found in worksheet " & ws.name
        Set searchHeaderByName = Nothing
    End If

End Function

'Function that returns true if a worksheet exist in a specified workbook.
Function sheetExistsByName(wb As Workbook, name As String) As Boolean
    exist = False
    For Each Sheet In wb.Worksheets
        If Sheet.name = name Then
            exist = True
        End If
    Next Sheet
    sheetExistsByName = exist
End Function

'Search for empty cells in the current selection and put "-" instead.
Sub cleanEmptyCells()
n_rows = Selection.Rows.Count
n_cols = Selection.Columns.Count
For i = 1 To n_rows
  For j = 1 To n_cols
    If IsEmpty(Selection.Cells(i, j).Value) Then
        Selection.Cells(i, j) = "-"
    End If
  Next j
Next i
End Sub

'Search for spaces and replaces with empty string, in the current selection.
Sub cleanStartWhiteSpaces()
n_rows = Selection.Rows.Count
n_cols = Selection.Columns.Count
For i = 1 To n_rows
  For j = 1 To n_cols
    Selection.Cells(i, j).Value = Replace(Selection.Cells(i, j).Value, " ", "")
  Next j
Next i
End Sub