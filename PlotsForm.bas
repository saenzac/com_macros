Private Sub BorrarButton_Click()
Dim i As Integer
For i = ListBoxRight.ListCount() - 1 To 0 Step -1
    If ListBoxRight.Selected(i) Then
        ListBoxRight.Selected(i) = False
        ListBoxRight.RemoveItem (i)
    End If
Next
End Sub

Private Sub CommandButton1_Click()
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Me.Hide
CalcularForm.Show
End Sub

Private Sub GraficarButton_Click()

ThisWorkbook.Worksheets("graficas").Select

Application.ScreenUpdating = False
ActiveWindow.DisplayGridlines = False

'remove old plots
If Worksheets("graficas").ChartObjects.Count > 0 Then
   Worksheets("graficas").ChartObjects.Delete
End If

Dim Final As Variant
Dim rangedata As Variant
Dim rangetime As Variant

Dim exp_name As String
Dim i As Integer
'MsgBox ListBoxRight.ListCount()
For i = 0 To ListBoxRight.ListCount - 1
    'MsgBox (ListBoxRight.List(i))
    exp_name = ListBoxRight.List(i)
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    Final = Sheets(exp_name).Cells(1, 1).End(xlDown).Row
    Set rangedata = Sheets(exp_name).Cells(2, 2).Resize(Final - 1, 2)
    Set rangetime = Sheets(exp_name).Cells(2, 1).Resize(Final - 1, 1)
    ActiveChart.SetSourceData Source:=rangedata
    ActiveChart.SeriesCollection(1).XValues = rangetime
    ActiveChart.SeriesCollection(1).name = "Input"
    ActiveChart.SeriesCollection(2).XValues = rangetime
    ActiveChart.SeriesCollection(2).name = "Output"
    ActiveChart.Location Where:=xlLocationAsObject, name:="graficas"
    ActiveChart.HasLegend = True    'Leyenda

    
    ActiveChart.HasDataTable = False   'Tiene Tabla de Datos
    
    With ActiveChart
            'Titulo Principal
            .HasTitle = True
            .ChartTitle.Characters.Text = "Experiment: " & exp_name
            'Titulo Horizontal
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time"
            'Titulo Vertical
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Voltage, Angle"
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
Next i

Application.ScreenUpdating = True
    
Call SizeGrafAct
'Sheets("graficas").Select

End Sub

Sub SizeGrafAct()
    Dim W As Double
    Dim H As Double
    Dim NumCols As Double
    
    Application.ScreenUpdating = False
    W = 476.25
    H = 241.5
    If ActiveSheet.ChartObjects.Count Mod 3 = 0 Then
        NumCols = Int(ActiveSheet.ChartObjects.Count / 3)
    Else
        NumCols = Int(ActiveSheet.ChartObjects.Count / 3) + 1
    End If
    
    Dim i As Integer
    For i = 1 To ActiveSheet.ChartObjects.Count
        With ActiveSheet.ChartObjects(i)
            .Width = W
            .Height = H
            .Left = 10 + ((i - 1) Mod NumCols) * W 'LeftPosition
            .Top = 10 + Int((i - 1) / NumCols) * H 'TopPosition
        End With
    Next i
    Application.ScreenUpdating = True
End Sub



Private Sub SendRightButton_Click()
Dim i As Integer
For i = 0 To ListBoxLeft.ListCount() - 1
    If ListBoxLeft.Selected(i) Then
        ListBoxRight.AddItem ListBoxLeft.List(i)
    End If
Next
End Sub

Private Sub UserForm_initialize()
Dim el As Worksheet
For Each el In CalcularForm.expsheets
    Me.ListBoxLeft.AddItem el.name
Next

End Sub