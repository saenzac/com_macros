''Module that contain simple functions to work in excel more efficiently

'Clear the whole contents of a sheet
'It asks for confirmation before proceed
Sub clearAllSheetContents()
    ans = MsgBox("Are you sure you want to clear all sheet contents?", vbYesNo)
    If ans = vbYes Then
        Cells.Select
        Selection.Clear
    End If
End Sub

'Set all sheet contents font to 8
Sub putAllSheetFont8()
    Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 8
    End With
End Sub

'Autofit all columns of the active sheet
Sub autofitAllColsOfSheet()
    Columns("A:A").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Cells.EntireColumn.AutoFit
End Sub

'Jump to a column(specified by letter) or to a row (specified by a  number)
Sub JumpTo()
Dim sResult As String
    On Error Resume Next
    
    sResult = InputBox("Type a row number or column letter and press Enter.", "Jump to...")
    
    If IsNumeric(sResult) Then
        Cells(sResult, ActiveCell.Column).Select
    Else
        Cells(ActiveCell.Row, sResult).Select
    End If

End Sub


'Function that Copy each worksheet to a new workbook, then
'calculate its weight (in bytes or mbytes)
'and make a summary listing all the sheets
'with their weights.
Sub worksheetsizes()
    Dim wks As Worksheet
    Dim c As Range
    Dim sFullFile As String
    Dim sReport As String
    Dim sWBName As String
    
    sReport = "Size Report"
    sWBName = "EraseMe.xls"
    sFullFile = ThisWorkbook.Path & _
    Application.PathSeparator & sWBName
    
    'Add new worksheet to record sizes
    On Error Resume Next
    Set wks = Worksheets(sReport)
    'If wks worksheet doesnt exits then create it
    If wks Is Nothing Then
        With ThisWorkbook.Worksheets.Add(Before:=Worksheets(1))
            .name = sReport
            .Range("A1").Value = "Worksheet Name"
            .Range("B1").Value = "Approximate Size"
        End With
    End If
    
    On Error GoTo 0
    With ThisWorkbook.Worksheets(sReport)
        .Select
        .Range("A1").CurrentRegion.Offset(1, 0).ClearContents
        Set c = .Range("A2")
    End With
    
    Application.ScreenUpdating = False
    'Loop through worksheets
    '1024 bytes = 1kb , fac = 1/1024
    Fact = 1 / 1024
    For Each wks In ActiveWorkbook.Worksheets
        If wks.name <> sReport Then
            wks.Copy
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveAs sFullFile
            ActiveWorkbook.Close savechanges:=False
            Application.DisplayAlerts = True
            c.Offset(0, 0).Value = wks.name
            c.Offset(0, 1).Value = FileLen(sFullFile) * Fact
            Set c = c.Offset(1, 0)
            Kill sFullFile
        End If
    Next wks
    Application.ScreenUpdating = True
End Sub
