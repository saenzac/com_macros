Dim csv_path As Variant

'Select a csv file and save the path, also the
'experiment name is selected
Private Sub Browse_Button_Click()
    Dim suggested_name As String
    
    csv_path = Application.GetOpenFilename("Excel files (*.csv), *.csv", , "Select a file", , False)
    If csv_path = False Then
        Exit Sub
    Else
        SelectedExp_textbox.Text = csv_path
    End If

    suggested_name = "exp" & ThisWorkbook.Worksheets.Count - 3 + 1
    Exp_name_textbox.Text = suggested_name
End Sub

'Copy the csv file contents to the selected experiment sheet.
Private Sub InsertarButton_Click()

    Dim workbook_csv As Workbook
    Dim newsheet As Worksheet
    Dim exp_name As String

    exp_name = Exp_name_textbox.Text
    
    If Not existSheetByName(exp_name) Then
        Set workbook_csv = Workbooks.Open(csv_path)
        Set newsheet = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newsheet.name = exp_name
    
        workbook_csv.Sheets(1).Cells.Copy ThisWorkbook.Worksheets(exp_name).Cells
        workbook_csv.Close SaveChanges:=False
        
        Range("A1").EntireRow.Insert shift:=xlDown
        
        With ThisWorkbook.Worksheets(exp_name)
            .Range("A1") = "Tiempo"
            .Range("B1") = "Entrada"
            .Range("C1") = "Salida"
        End With
        
        Range("A1").EntireRow.Font.Bold = True

        CalcularForm.expsheets.Add Item:=newsheet
        CalcularForm.Select_experiment_textbox.AddItem newsheet.name

        Me.Hide
        CalcularForm.Show
    Else
        MsgBox "Ya existe una hoja con el nombre dado, especificar otro nombre", Title:="Error"
    End If
End Sub

Function existSheetByName(name As String) As Boolean
    Dim exists As Boolean, el As Object
    exists = False
    For Each el In ThisWorkbook.Worksheets
        If el.name = name Then
            exists = True
        End If
    Next
    existSheetByName = exists
End Function

Private Sub VolverButton_Click()
Me.Hide
CalcularForm.Show
End Sub