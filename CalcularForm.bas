Public mainsheet As Object
Public inicell As Object
Public inirow As Integer
Public inicol As Integer
Public costo_cell As Object
Public expsheets As Collection

Private Sub CerrarButton_Click()
Unload Me
End Sub


Private Sub UserForm_initialize()
    Dim expsheets() As Worksheet
    Set mainsheet = ThisWorkbook.Worksheets("main")
    Set inicell = mainsheet.Range("refcel")
    inirow = inicell.Row
    inicol = inicell.Column
    Set costo_cell = inicell.Offset(12, 6)

    Me.updateExpSheets
    
    Select_experiment_textbox.Clear
    Dim el As Worksheet
    For Each el In Me.expsheets
        Select_experiment_textbox.AddItem el.name
    Next

    a1_TB.Value = 0.1
    inicell.Offset(1, 1) = a1_TB.Value
    a2_TB.Value = 0.1
    inicell.Offset(1, 2) = a2_TB.Value
    a3_TB.Value = 0.1
    inicell.Offset(1, 3) = a3_TB.Value

End Sub

Private Sub PlotsButton_Click()
Me.Hide
PlotsForm.Show
End Sub




Private Sub a1_TB_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(a1_TB.Value) Then
    inicell.Offset(1, 1) = a1_TB.Value
    'costo_TB.Value = Format(costo_cell.Value, "0.000")
    costo_TB.Value = 0#
    Else
    MsgBox "Ingresar un valor numerico"
    a1_TB.Value = 0.1
    End If
End Sub

Private Sub a2_TB_Change()
    If IsNumeric(a2_TB.Value) Then
    inicell.Offset(1, 2) = a2_TB.Value
    'costo_TB.Value = Format(costo_cell.Value, "0.000")
    costo_TB.Value = 0#
    Else
    MsgBox "Ingresar un valor numerico"
    a2_TB.Value = 0.1
    End If
End Sub

Private Sub a3_TB_Change()
    If IsNumeric(a3_TB.Value) And a3_TB.Value <> 0 Then
    inicell.Offset(1, 3) = a3_TB.Value
    'costo_TB.Value = Format(costo_cell.Value, "0.000")
    costo_TB.Value = 0#
    Else
    MsgBox "Ingresar un valor numerico diferente de 0"
    a3_TB.Value = 0.1
    End If
End Sub

Private Sub EstimarButton_Click()
    If Select_experiment_textbox.ListIndex = -1 Then
        MsgBox "Seleccionar algun experimento de la lista"
        Exit Sub
    End If
    'Data preparation
    Dim exp_name As String
    Dim expsheet As Worksheet
    Dim destcell As Object
    Dim originrange As Object
    Dim error_ini_cell As Object
    Dim error_end_cell As Object
    Dim error_ini_address As String
    Dim error_end_address As String
    Dim salmodelo_ini_cell As Object
    Dim salmodelo_end_cell As Object
    Dim costo_cell_address As String
    Dim param_range_address As String
    Dim last_row_exp As Integer
    Dim last_row_main As Integer
    Dim salmodelo_ini_address As Variant
    Dim salmodelo_end_address As Variant

    mainsheet.Select

    'Create the experiment - selected worksheet object
    exp_name = Select_experiment_textbox.SelText
    Set expsheet = Worksheets(exp_name)

    'Copy cells from the experiment sheet to the main sheet
    last_row_exp = expsheet.Cells(Rows.Count, 1).End(xlUp).Row

    Set destcell = inicell.Offset(12, 0)
    Set originrange = expsheet.Range("A2:C" & last_row_exp)

    originrange.Copy Destination:=destcell

    last_row_main = mainsheet.Cells(Rows.Count, inicol).End(xlUp).Row

    Set error_ini_cell = inicell.Offset(12, 5)
    Set error_end_cell = Cells(last_row_main, error_ini_cell.Column)
    Set salmodelo_ini_cell = inicell.Offset(12, 4)
    Set salmodelo_end_cell = Cells(last_row_main, salmodelo_ini_cell.Column)
    
    error_ini_address = error_ini_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    error_end_address = error_end_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    salmodelo_ini_address = salmodelo_ini_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    salmodelo_end_address = salmodelo_end_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'T[n] = a*V[n-3] + b*T[n-3] + c*T[n-2] + d*T[n-1]
    Dim cell As Variant
    For Each cell In mainsheet.Range(salmodelo_ini_address & ":" & salmodelo_end_address)
        'Formula de 'Salida de modelo'
        cell.FormulaR1C1 = "=a.*R[-3]C[-3] - b.*R[-3]C[0] - c.*R[-2]C[0] - d.*R[-1]C[0]"
        'Formula de 'Error'
        cell.Offset(0, 1).FormulaR1C1 = "= R[0]C[-3] - R[0]C[-1]"
    Next

    costo_cell.Formula = "=SUMSQ(" & error_ini_address & ":" & error_end_address & ")"

    'Calculate best model
    costo_cell_address = costo_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    param_range_address = inicell.Offset(1, 1).Resize(columnsize:=3).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Call SolverReset
    SolverOptions precision:=0.001
    SolverAdd CellRef:=param_range_address, Relation:=3, FormulaText:="0.01"
    SolverOk SetCell:=costo_cell_address, _
             MaxMinVal:=2, _
             ValueOf:="0", _
             ByChange:=param_range_address, _
             EngineDesc:="GRG Nonlinear"
    SolverSolve userFinish:=True
    SolverFinish KeepFinal:=1

     a1_TB.Value = Format(inicell.Offset(1, 1).Value, "0.000")
     a2_TB.Value = Format(inicell.Offset(1, 2).Value, "0.000")
     a3_TB.Value = Format(inicell.Offset(1, 3).Value, "0.000")
    
End Sub

Private Sub InsNuevoExpButton_click()
    Me.Hide
    InsertarExp.Show
End Sub

Private Function deleteExpSheets()
    Set expsheets = Nothing
    Set expsheets = New Collection
End Function

Public Function updateExpSheets()
    deleteExpSheets
    Dim el As Worksheet
    For Each el In ThisWorkbook.Worksheets
        If el.name <> "main" And el.name <> "graficas" And el.name <> "validacion" Then
            expsheets.Add Item:=el
        End If
    Next
End Function

Private Sub ValidarButton_Click()
    mainsheet.Select

    ThisWorkbook.Worksheets("validacion").Select

    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False

    'remove old plots
    If Worksheets("validacion").ChartObjects.Count > 0 Then
       Worksheets("validacion").ChartObjects.Delete
    End If

    Dim rangerealoutput As Object
    Dim rangeestoutput As Object
    Dim rangetime As Object

    last_row_main = mainsheet.Cells(Rows.Count, inicol).End(xlUp).Row

    Set time_ini_cell = inicell.Offset(12, 0)
    Set time_end_cell = Cells(last_row_main, time_ini_cell.Column)
    Set sal_ini_cell = inicell.Offset(12, 2)
    Set sal_end_cell = Cells(last_row_main, sal_ini_cell.Column)
    Set salmodelo_ini_cell = inicell.Offset(12, 4)
    Set salmodelo_end_cell = Cells(last_row_main, salmodelo_ini_cell.Column)

    time_ini_address = time_ini_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    time_end_address = time_end_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    sal_ini_address = sal_ini_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    sal_end_address = sal_end_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    salmodelo_ini_address = salmodelo_ini_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    salmodelo_end_address = salmodelo_end_cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Set rangerealoutput = mainsheet.Range(sal_ini_address & ":" & sal_end_address)
    Set rangeestoutput = mainsheet.Range(salmodelo_ini_address & ":" & salmodelo_end_address)
    Set rangetime = mainsheet.Range(time_ini_address & ":" & time_end_address)

    ThisWorkbook.Worksheets("validacion").Select

    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine

    ActiveChart.SeriesCollection.Add Source:=rangerealoutput
    ActiveChart.SeriesCollection.Add Source:=rangeestoutput
    ActiveChart.SeriesCollection(1).name = "Salida medida"
    ActiveChart.SeriesCollection(1).XValues = rangetime
    ActiveChart.SeriesCollection(2).name = "Salida estimada"
    ActiveChart.SeriesCollection(2).XValues = rangetime
    ActiveChart.SeriesCollection(2).Format.Line.DashStyle = msoLineSysDash
    
    ActiveChart.Location Where:=xlLocationAsObject, name:="validacion"
    ActiveChart.HasLegend = True    'Leyenda

    ActiveChart.HasDataTable = False   'Tiene Tabla de Datos

    With ActiveChart.Parent
        .Top = Range("B2").Top
        .Left = Range("B2").Left
        .Width = Range("B2:L2").Width
        .Height = Range("B2:B25").Height
    End With
    
    
    With ActiveChart
            'Titulo Principal
            .HasTitle = True
            .ChartTitle.Characters.Text = "Experiment: " & exp_name
            'Titulo Horizontal
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
            'Titulo Vertical
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Angle"
    End With
    
    'Atributos del Título principal
    ActiveChart.ChartTitle.Select
    Selection.AutoScaleFont = True
    With Selection.Font
            .name = "Arial"
            .Size = 10
            .Bold = True
    End With

         'Atributos del Título vertical
        ActiveChart.Axes(xlValue).AxisTitle.Select
        Selection.AutoScaleFont = True
        With Selection.Font
            .name = "Arial"
            .Size = 8
            .Bold = True
        End With
        
        'Atributos del Título horizontal
        ActiveChart.Axes(xlCategory).AxisTitle.Select
        Selection.AutoScaleFont = True
        With Selection.Font
            .name = "Arial"
            .Size = 8
            .Bold = True
        End With
        
        'Borde del Gráfico
        With Selection.Border
            .Weight = 4
            .LineStyle = -1
        End With


Application.ScreenUpdating = True
End Sub