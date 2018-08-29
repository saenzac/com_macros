'Function that updates all the pivot tables in the active sheet
'using an index number which represents
'the data source (range object which contains the data)
Sub updatePivotTables()
    Dim pt_rng As Range
    Dim ind As Integer
    Dim newrange As String
 
    'show how many elements contains pivotcaches collection
    MsgBox "This workbook already contains " & ActiveWorkbook.PivotCaches.Count & " pivot caches"

    For Each PT In ActiveSheet.PivotTables
        Set pt_rng = PT.TableRange2  'tablerange2 select including pivot table pagefields
        ind = Cells(2, pt_rng.Column).Value

        newrange = getRangeByIndex(ind) 'see FunctionsHelper
        PT.ChangePivotCache ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newrange)
        PT.RefreshTable
    Next PT

    MsgBox "Pivot tables succesfully udpated"
End Sub


'Function that from an origin worksheet copy specified headers to the destination worksheet.
'The headers to copy are specified in the "params" worksheet.
Sub copyColumnsByHeaderName()
    Dim ws_origin As Worksheet
    Dim ws_dest As Worksheet
    Dim ws_params As Worksheet
    Dim ws_name_origin As String
    Dim ws_name_dest As String

    Set ws_params = ActiveWorkbook.ActiveSheet

    ws_name_origin = ws_params.Range("B2")
    ws_name_dest = ws_params.Range("D2")
    
    If sheetExistsByName(ActiveWorkbook, ws_name_dest) Then
        ans = MsgBox("Ya existe un libro con el nombre: " & ws_name_dest & "desea re-escribirlo?", vbYesNo)
        If ans = vbNo Then
            Exit Sub
        End If
        Set ws_dest = Worksheets(ws_name_dest)
    Else
        Set ws_dest = addWorkSheetByName(ActiveWorkbook, ws_name_dest)
    End If

    If sheetExistsByName(ActiveWorkbook, ws_name_origin) Then
        Set ws_origin = ActiveWorkbook.Worksheets(ws_name_origin)
    Else
        MsgBox "Hoja " & ws_name & " no encontrada "
        Exit Sub
    End If

    finalcol = ws_params.Cells(5, ws_params.Columns.Count).End(xlToLeft).Column

    For i = 2 To finalcol
        breakfor = False

        ws_params.Activate
        ws_dest.Cells(1, i - 1) = ws_params.Cells(5, i).Value

        If ws_params.Cells(4, i) = "-" Then
            breakfor = True
        End If

        If Not breakfor Then
            Dim rangef As Range
            Set rangef = searchHeaderByName(ws_origin, ws_params.Cells(4, i))
            If rangef Is Nothing Then
                MsgBox "Error hallando cabecera  ... Finalizando funcion"
                Exit Sub
            Else
                ws_origin.Activate
                rangef.Offset(1, 0).Select
                Range(Selection, Selection.Offset(100000, 0)).Select
                Selection.Copy
                ws_dest.Activate
                Cells(1 + 1, i - 1).Select
                Selection.PasteSpecial Paste:=xlValues
            End If
        End If
    Next i
    
    ws_dest.Activate
    Call putAllSheetFont8
    Call autofitAllColsOfSheet
End Sub



'Find rows delimiters of the regions that compose the "cuotas captura" sheet.
Sub getCuotaCapturaRanges()

    Dim wscc As Worksheet
    Set wscc = ActiveWorkbook.Worksheets("cc")
    Dim wscuotacaptura As Worksheet
    Set wscuotacaptura = ActiveWorkbook.Worksheets("Cuotas Captura")
    
    wscuotacaptura.Select
    
    Range("A1").Select
    Dim r_inicio As Long

    i = 1
    Do While True
        
        Selection.End(xlDown).Select
        r_inicio = Selection.Row
        If r_inicio > 10000 Then
            Exit Do
        End If
    
        wscc.Cells(i, 1) = Selection.Value
        Selection.End(xlDown).Select
        r_final = Selection.Row
        wscc.Cells(i, 2) = r_inicio
        wscc.Cells(i, 3) = r_final
        i = i + 1
    Loop

    wscc.Select

End Sub


'Read the actual visibility settings of the sheets and show that status
Sub readSheetsStatus()

Dim ws_sheet  As Worksheet
Set ws_sheet = Worksheets("libros")

ws_sheet.Activate

i = 5
For Each sh In ActiveWorkbook.Worksheets
    If sh.name <> "libros" Then
        Cells(i, 1) = sh.name
        If sh.Visible Then
            Cells(i, 2) = 1
        Else
            Cells(i, 2) = 0
        End If
        i = i + 1
    End If
Next sh

End Sub

'Set the custom visibility of the sheets
Sub setSheetsStatus()

Dim ws_sheet  As Worksheet
Set ws_sheet = Worksheets("libros")

ws_sheet.Activate

ultirow = ws_sheet.Cells(ws_sheet.Rows.Count, 1).End(xlUp).Row

Dim name As String
For i = 5 To ultirow
    name = Cells(i, 1)
    If Cells(i, 2) <> "libros" Then
        If Cells(i, 2) = 1 Then
            ActiveWorkbook.Worksheets(name).Visible = True
        Else
            ActiveWorkbook.Worksheets(name).Visible = xlSheetHidden
        End If
    End If
Next i

End Sub


'Function that creates pivot tables
'according to a small table where parameters
'as rows, columns anda data values which it will contain
'are specified
Sub createPivotTablesFromParams()
    Dim AWB As Workbook
    Set AWB = ActiveWorkbook

    Dim WSDT As Worksheet
    Dim WSOrigin As Worksheet
    Dim WSDScheme As Worksheet
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim PRange As String
    Dim FinalRow As Long

    'Book where the dynamic table (dt) is created
    Set WSDT = Worksheets("Pivot Inar Consultor2")
    'Book origin of the data
    Set WSOrigin = Worksheets("Inar Total")
    'Book parameters origin for the dt
    Set WSDScheme = Worksheets("pivotsdef")

    'Clean up the dynamic tables
    For Each PT In WSDT.PivotTables
        PT.TableRange2.Clear
    Next PT

    'We select data range
    FinalRow = WSOrigin.Cells(Rows.Count, 1).End(xlUp).Row
     PRange = getRangeByIndex(1)
    colDT = 2 'column number of the pivot table
    rowDT = 10 'row number of the pivot table
    r = 3 'numero de fila tabla de parametros
    c = 2 'numero de columna tabla de parametros
    j = 1 'numeracion para diferencia nombre de las td

    Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange)
    Do While WSDScheme.Cells(3, c) <> ""
        WSDT.Select
        Set PT = PTCache.CreatePivotTable(TableDestination:=WSDT.Cells(rowDT, colDT), TableName:="pivotable1" & j)
        'PT.Format PivotStyleLight4
        PT.ManualUpdate = True

        WSDScheme.Select
        'Iteramos por los valores de una columna agregando las cadenas al pivot
        With PT
            nombre_pivot = Cells(r, c)
            WSDT.Cells(1, colDT) = nombre_pivot
            r = r + 1
        
            'agrega pivot data fields
            Do While Cells(r, c).Value <> ""
                cad = Cells(r, c)
                .AddDataField Field:=.PivotFields(cad), Function:=xlSum
                r = r + 1
            Loop
        
            '.AddDataField .PivotFields("Gross"), "Sum of Profit", xlSum
        
            Do While Cells(r, c).Value = ""
                r = r + 1
            Loop
        
            'agrega pivot filters
            Do While Cells(r, c).Value <> ""
              cad = Cells(r, c)
              If cad = "No" Then
                r = r + 1
                Exit Do
              End If
        
              With .PivotFields(cad)
                .Orientation = xlPageField 'inserta el filtro, lo hace visible
                
                'For i = 1 To .PivotItems.Count
                '    If InStr(1, .PivotItems(i), "2G") <> 0 Then
                '        .PivotItems(i).Visible = True
                '    Else
                '        .PivotItems(i).Visible = False
                '    End If
                'Next i
              End With
              r = r + 1
            Loop
        
            Do While Cells(r, c).Value = ""
                r = r + 1
            Loop
        
            'agrega pivol cols
            Do While Cells(r, c).Value <> ""
                cad = Cells(r, c)
                If cad = "No" Then
                  r = r + 1
                  Exit Do
                End If
        
                With .PivotFields(cad)
                    .Orientation = xlColumnField
                End With
                r = r + 1
            Loop
        
            Do While Cells(r, c).Value = ""
                r = r + 1
            Loop
        
            'agrega pivot rows
            Do While Cells(r, c).Value <> ""
                cad = Cells(r, c)
                With .PivotFields(cad)
                    .Orientation = xlRowField
                End With
                r = r + 1
            Loop
            
        End With
        
        PT.ManualUpdate = False
        
        ancho = PT.DataBodyRange.Columns.Count
        'pfields = PT.PageFields.Count
        'rowscount = PT.DataLabelRange.Columns.Count
        totalcols = ancho + rowscount
        colDT = colDT + totalcols + 1 + 1
        j = j + 1
        c = c + 2
        r = 3
    Loop
    
    WSDT.Select
    Cells.Select
    With Selection.Font
        .name = "Calibri"
        .Size = 7
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
End Sub
    
    
'mode:  ruc, noruc, admin
Sub genSummaryAsIntegral()
        mode = "asvent_noruc"
        'Open "as. integral" main file
        paramsbookfile = "params_resumen.xlsm"
        Dim f_asint As Workbook
        Dim f_asventas As Workbook
        Set f_asint = getWorkbook(mainpath & "\tiendas\Comisiones_Asesores_Int. Nuevo Esquema Julio 2018.xlsm")
        Set f_asventas = getWorkbook(mainpath & "\tiendas\Comisiones_Caps y Nuevo esquema Julio 2018.xlsm")
    
        'Create a new blank workbook, and assign it a name depending on the mode
        Dim f_res_wb As Workbook
        Set f_res_wb = Workbooks.Add
    
        If mode = "ruc" Then
            f_res_wb.SaveAs Filename:=mainpath & "\resumen\resumen_ruc.xlsx"
            paramsbookname = "ResumenRUC"
        ElseIf mode = "asint_noruc" Then
            f_res_wb.SaveAs Filename:=mainpath & "\resumen\resumen_as_integral.xlsx"
            paramsbookname = "ResumenASINT"
        ElseIf mode = "asvent_noruc" Then
            f_res_wb.SaveAs Filename:=mainpath & "\resumen\resumen_as_ventas.xlsx"
            paramsbookname = "ResumenASVENTAS"
        End If
        
    
        Dim paramsSheet As Worksheet
        Set paramsSheet = Workbooks(paramsbookfile).Worksheets(paramsbookname)
        lastrow = paramsSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 3 To lastrow
            file = paramsSheet.Cells(i, 1)
            Dim bookor As String
            bookor = paramsSheet.Cells(i, 2)
            bookdest = paramsSheet.Cells(i, 3)
    
            If file = "ASINT" Then
                If Not existBookInWB(f_asint.Worksheets, bookor) Then
                    MsgBox "The book '" & bookor & "' doesn't exist"
                    Exit For
                End If
    
                sheetcount = f_res_wb.Worksheets.Count
                f_asint.Sheets(bookor).Copy after:=f_res_wb.Sheets(sheetcount)
                f_res_wb.Worksheets(bookor).Visible = True
                f_res_wb.Worksheets(bookor).Select
                f_res_wb.Worksheets(bookor).name = bookdest
                f_res_wb.Worksheets(bookdest).Cells.Select
                Application.CutCopyMode = False
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ElseIf file = "ASVENT" Then
                If Not existBookInWB(f_asventas.Worksheets, bookor) Then
                    MsgBox "The book " & bookor & " doesn't exist"
                    Exit For
                End If
    
                sheetcount = f_res_wb.Worksheets.Count
                f_asventas.Sheets(bookor).Copy after:=f_res_wb.Sheets(sheetcount)
                f_res_wb.Worksheets(bookor).Visible = True
                f_res_wb.Worksheets(bookor).Select
                f_res_wb.Worksheets(bookor).name = bookdest
                f_res_wb.Worksheets(bookdest).Cells.Select
                Application.CutCopyMode = False
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            End If
        Next i
End Sub


Function alcanceb(area As Range, bafi As Range)
    If bafi.Value = "BAFI" Then
        If area.Value = "POST VENTA TP" Then
            alcance = "ASESOR POST VENTA - SIN BAFI"
        ElseIf area.Value = "BIENVENIDA TP" Then
            alcance = "ASESOR DE BIENVENIDA - SIN BAFI"
        ElseIf area.Value = "COORDINADOR TP" Then
            alcance = "COORDINADOR DE PISO - SIN BAFI"
        End If
    Else
        If area.Value = "POST VENTA TP" Then
            alcance = "ASESOR POST VENTA"
        ElseIf area.Value = "BIENVENIDA TP" Then
            alcance = "ASESOR DE BIENVENIDA"
        ElseIf area.Value = "COORDINADOR TP" Then
            alcance = "COORDINADOR DE PISO"
        End If
    End If
End Function


Sub detectBlankAtStartOrEnd()
    str = Selection.Cells(1, 1)
    r = Right(str)
  '  l = Left(str, 1)
    MsgBox code(r)
End Sub

