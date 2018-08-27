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