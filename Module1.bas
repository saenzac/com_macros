Sub removeAllSheetsExceptMain()
For Each hoja In ThisWorkbook.Worksheets
    If hoja.name <> "main" And hoja.name <> "validacion" And hoja.name <> "graficas" Then
        hoja.Delete
    End If
Next
End Sub

Sub test()
ActiveWindow.DisplayGridlines = True
End Sub