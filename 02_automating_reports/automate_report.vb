Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1


Public Sub formulario()
    
    
    
End Sub

Public rango_defecto, clasif_puntos As Boolean
Public cantidad_activos As Integer

Private Sub CommandButton1_Click()
       
    Dim cual_tabla As Integer
    cual_tabla = 6
    
    Call DesCombinar
    
    If OptionButton3.Value = True Then
        'MsgBox "Por Activo"
        Call EliminarColumnaVehiculos
        Call ClasificacionPorPuntosColor(TextBox1.Value, TextBox2.Value, TextBox3.Value, TextBox4.Value, TextBox5.Value, TextBox6.Value, cual_tabla, "V")
        Call titulos_verticales("V5")
        Call expandirActivo
        Call titulo("V")
        Call Datos_graficaActivo
        Call tratamiento_tabla2("Graficas Activo")
        Call diseño_tabla2
        Worksheets("Graficas Activo").Activate
        cual_tabla = 2
        Call ClasificacionPorPuntosColor(TextBox1.Value, TextBox2.Value, TextBox3.Value, TextBox4.Value, TextBox5.Value, TextBox6.Value, cual_tabla, "V")
        Call graficar("Graficas Activo")
        
    Else
        'MsgBox "Por Operador"
        Call EliminarColumnaOperador
        Call ClasificacionPorPuntosColor(TextBox1.Value, TextBox2.Value, TextBox3.Value, TextBox4.Value, TextBox5.Value, TextBox6.Value, cual_tabla, "Z")
        Call titulos_verticales("Z5")
        Call expandirOperador
        Call titulo("Z")
        Call Datos_graficaOperador
        Call tratamiento_tabla2("Graficas Operador")
        Call diseño_tabla2
        Worksheets("Graficas Operador").Activate
        cual_tabla = 2
        Call ClasificacionPorPuntosColor(TextBox1.Value, TextBox2.Value, TextBox3.Value, TextBox4.Value, TextBox5.Value, TextBox6.Value, cual_tabla, "V")
        Call graficar("Graficas Operador")

        
    End If
    

    
    'MsgBox "Activos = " & TextBox7.Value



    Unload UserForm1

End Sub

Private Sub CommandButton2_Click()
    'Application.Quit
    Unload UserForm1
    'Exit Sub
End Sub

Private Sub CommandButton3_Click()
    
    If TextBox1.Value = "" Then
    
        CommandButton1.Enabled = False
    
    End If
    
    If TextBox1.Value <> "" And TextBox2.Value <> "" And TextBox3.Value <> "" And TextBox4.Value <> "" And TextBox5.Value <> "" And TextBox6.Value <> "" Then
    
        CommandButton1.Enabled = True
            
    End If
    

End Sub



Private Sub OptionButton1_Click()
    
    
    TextBox1.Enabled = False
    TextBox2.Enabled = False
    TextBox3.Enabled = False
    TextBox4.Enabled = False
    TextBox5.Enabled = False
    TextBox6.Enabled = False
    TextBox1.Value = 0
    TextBox2.Value = 79
    TextBox3.Value = 80
    TextBox4.Value = 89
    TextBox5.Value = 90
    TextBox6.Value = 100
    
    
    CommandButton1.Enabled = True
    
    rango_defecto = True
        
    
End Sub

Private Sub OptionButton2_Click()
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    TextBox3.Enabled = True
    TextBox4.Enabled = True
    TextBox5.Enabled = True
    TextBox6.Enabled = True
    CommandButton1.Enabled = False
    rango_defecto = False
End Sub
Private Sub DesCombinar()
    'descombinar celdas A, B, C y D
    Range("A5").Select 'Siempre llega en este formato
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
End Sub

Private Sub EliminarColumnaVehiculos() 'ESTO SE PUEDE CAMBIAR POR UN ALGORITMO DE BUSQUEDA DE PARAMETROS Y ELIMINACIÓN!!!
    Range("AI:AI").Select
    Selection.EntireColumn.Delete
 
    Range("AH:AH").Select
    Selection.EntireColumn.Delete
 
    Range("AG:AG").Select
    Selection.EntireColumn.Delete
 
    Range("AF:AF").Select
    Selection.EntireColumn.Delete
 
    Range("AE:AE").Select
    Selection.EntireColumn.Delete
 
    Range("AD:AD").Select
    Selection.EntireColumn.Delete
 
    Range("AC:AC").Select
    Selection.EntireColumn.Delete
 
    Range("AA:AA").Select
    Selection.EntireColumn.Delete
 
    Range("Z:Z").Select
    Selection.EntireColumn.Delete
 
    Range("X:X").Select
    Selection.EntireColumn.Delete
 
    Range("W:W").Select
    Selection.EntireColumn.Delete
 
    Range("U:U").Select
    Selection.EntireColumn.Delete
 
    Range("T:T").Select
    Selection.EntireColumn.Delete
 
    Range("D:D").Select
    Selection.EntireColumn.Delete
 
    Range("W:W").Select
    Selection.EntireColumn.Delete
End Sub

Private Sub EliminarColumnaOperador() 'ESTO SE PUEDE CAMBIAR POR UN ALGORITMO DE BUSQUEDA DE PARAMETROS Y ELIMINACIÓN!!!

    Range("AL:AL").Select
    Selection.EntireColumn.Delete
 
    Range("AK:AK").Select
    Selection.EntireColumn.Delete
 
    Range("AJ:AJ").Select
    Selection.EntireColumn.Delete
 
    Range("AI:AI").Select
    Selection.EntireColumn.Delete
 
    Range("AH:AH").Select
    Selection.EntireColumn.Delete
 
    Range("AG:AG").Select
    Selection.EntireColumn.Delete
 
    Range("AF:AF").Select
    Selection.EntireColumn.Delete
 
    Range("AD:AD").Select
    Selection.EntireColumn.Delete
 
    Range("AC:AC").Select
    Selection.EntireColumn.Delete
 
    Range("AA:AA").Select
    Selection.EntireColumn.Delete
 
    Range("Z:Z").Select
    Selection.EntireColumn.Delete
 
    Range("X:X").Select
    Selection.EntireColumn.Delete
 
    Range("W:W").Select
    Selection.EntireColumn.Delete
    
    Range("AA:AA").Select
    Selection.EntireColumn.Delete
 
 

 End Sub

Private Sub ClasificacionPorIndicesColor()

    Dim i, cantidadFilas, x As Double
    Dim n, cell_V, cell_A, rango As String
    
    cantidad_activos = 0
    
    x = 6
    n = CStr(x)
    cell_V = "V" + n
    cell_A = "A" + n
    rango = cell_A + ":" + cell_V

    cantidadFilas = Contador_de_filas
    
    'MsgBox rango


    For i = 1 To cantidadFilas
    
        'MsgBox "Entró " & " i= " & i & " " & rango
        'If Range(cell).Value = 50 Then
        
        If Range(cell_V).Value >= 10 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(255, 199, 206) 'ROJO
        
        End If
        
        If Range(cell_V).Value >= 5.1 And Range(cell_V).Value <= 9.9 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(255, 235, 156) 'AMARILLO
        
        End If
        
        If Range(cell_V).Value >= 0 And Range(cell_V).Value <= 5 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(198, 239, 206) 'VERDE
        
        End If
        
        If Range(cell_A).Value = "SUB-TOTAL" Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(242, 242, 242) 'GRIS
            cantidad_activos = cantidad_activos + 1
        
        End If
        
        x = x + 1
        n = CStr(x)
        cell_V = "V" + n
        cell_A = "A" + n
        rango = cell_A + ":" + cell_V

    
    Next i


End Sub



Private Sub ClasificacionPorPuntosColor(ByVal l1 As Double, ByVal l2 As Double, ByVal l3 As Double, ByVal l4 As Double, ByVal l5 As Double, ByVal l6 As Double, ByVal cual_tabla As Integer, ByVal column As String)

    Dim i, cantidadFilas, x As Integer
    Dim n, cell_VW, cell_A, cell_T, cell, rango As String
    
    cantidad_activos = 0
    
    x = cual_tabla '6 'QUE ESTE VALOR SE PASE POR REFERENCIA!!!!!!
    n = CStr(x)
    cell_A = "A" + n

    If cual_tabla = 6 Then
        cell = column + n
        rango = cell_A + ":" + cell
    End If
    If cual_tabla = 2 Then
        cell = "T" + n
        rango = cell_A + ":" + cell
    End If

    cantidadFilas = Contador_de_filas(cual_tabla) 'AL VALOR POR REFERENCIA TAMBIEN ENVIARLE POR REFENCIA DESDE DONDE EMPEZAR A CONTAR
    
    'MsgBox l1
    'MsgBox cantidadFilas

    If rango_defecto = True Then
        
        l1 = 0
        l2 = 79
        l3 = 80
        l4 = 89
        l5 = 90
        l6 = 100
        
    End If

    For i = 1 To cantidadFilas
    
        'MsgBox "Entró " & " i= " & i & " " & rango
        'If Range(cell).Value = 50 Then
        
        If Range(cell).Value >= l1 And Range(cell).Value <= l2 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(255, 199, 206) 'ROJO
        
        End If
        
        If Range(cell).Value >= l3 And Range(cell).Value <= l4 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(255, 235, 156) 'AMARILLO
        
        End If
        
        If Range(cell).Value >= l5 And Range(cell).Value <= l6 Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(198, 239, 206) 'VERDE
        
        End If
        
        If Range(cell_A).Value = "SUB-TOTAL" Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(242, 242, 242) 'GRIS
            cantidad_activos = cantidad_activos + 1
        
        End If
        
        x = x + 1
        n = CStr(x)
        'cell_V = "V" + n
        cell_A = "A" + n
        If cual_tabla = 6 Then
            cell = column + n
        
        End If
        If cual_tabla = 2 Then
           cell = "T" + n
        End If
        rango = cell_A + ":" + cell

        If Range(cell_A).Value = "Total" Then
            
            'MsgBox "Entró"
            Range(rango).Interior.Color = RGB(242, 242, 242) 'GRIS
            'cantidad_activos = cantidad_activos + 1
        
        End If

    
    Next i
    
    'MsgBox cell_A
    If cual_tabla = 6 Then
        TextBox7.Value = cantidad_activos
    End If
    
    'For i = 1 To cantidadFilas
    'Next i

End Sub



Function Contador_de_filas(ByVal cual_tabla As Integer) As Integer ' Contador de filas desde V6
    
    Dim n, cell, cell_V, cell_T As String
    Dim flag As Boolean
    Dim x As Integer
    
    x = cual_tabla '6
    n = CStr(x)
    
        
    If cual_tabla = 6 Then
        cell = "A" + n
    End If
    If cual_tabla = 2 Then
        cell = "A" + n
    End If
    
    flag = True
    
    Do While flag
        If Range(cell).Value <> "" And Range(cell).Value <> "Total" Then
            x = x + 1
            n = CStr(x)
            If cual_tabla = 6 Then
                cell = "A" + n
            End If
            If cual_tabla = 2 Then
                cell = "A" + n
            End If
        Else
            flag = False
            x = x - 1
            n = CStr(x)
            'MsgBox "Cells: " + n
        End If
    Loop
    If cual_tabla = 6 Then
        Contador_de_filas = x - 5
    End If
    If cual_tabla = 2 Then
        Contador_de_filas = x - 1
    End If
End Function



Private Sub expandirActivo() 'ESTO SE PUEDE CAMBIAR POR UN ALGORITMO DE BUSQUEDA, ASI NO SE REPITE CÓDIGO Y SE REDUCE HACIENDOLO MÁS RÁPIDO!!!
    
    Range("H5").Value = "Duración"
    Range("K5").Value = "Duración"
    Range("N5").Value = "Duración"
    Range("Q5").Value = "Duración"
    Range("U5").Value = "Total Faltas"
    Range("V5").Value = "Total Calificación"
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Columns("S:S").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit

End Sub

Private Sub expandirOperador() 'ESTO SE PUEDE CAMBIAR POR UN ALGORITMO DE BUSQUEDA, ASI NO SE REPITE CÓDIGO Y SE REDUCE HACIENDOLO MÁS RÁPIDO!!!
    
    Range("D5").Value = "Compañia Conductor" 'ok
    Range("E5").Value = "Compañia Vehículo"  'ok
    'Range("J5").Value = "Duración"
    Range("L5").Value = "Duración"
    'Range("M5").Value = "Duración"
    Range("O5").Value = "Duración"
    'Range("P5").Value = "Duración"
    Range("R5").Value = "Duración"
    'Range("S5").Value = "Duración"
    Range("U5").Value = "Duración"
    Range("Y5").Value = "Total Faltas"
    Range("Z5").Value = "Total Calificación"
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Columns("S:S").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Columns("W:W").EntireColumn.AutoFit
    Columns("X:X").EntireColumn.AutoFit
    Columns("Y:Y").EntireColumn.AutoFit
    Columns("Z:Z").EntireColumn.AutoFit

End Sub

Private Sub titulos_verticales(ByVal cell As String)
    cell = "A5:" + cell
    Range(cell).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("5:5").EntireRow.AutoFit
End Sub

Private Sub titulo(ByVal column As String)
    
    Dim rango, rango2 As String
    rango = "A1:" + column + "4"
    rango2 = column + "1"
    Range(rango).Select
    Range(rango2).Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    
    If OptionButton3.Value = True Then
    
        'MsgBox "Por Activo"
        Range("A1").Activate
        With Selection
             .Value = "INFORME DETALLADO POR ACTIVO"
            .Font.Bold = True
            .Font.Size = 12
        End With
    
        ActiveSheet.Name = "Activo"
        Worksheets("Hoja2").Name = "Graficas Activo"
        Worksheets("Graficas Activo").Activate

    Else
        'MsgBox "Por Operador"
        Range("A1").Activate
        With Selection
             .Value = "INFORME DETALLADO POR OPERADOR"
            .Font.Bold = True
            .Font.Size = 12
        End With
    
        ActiveSheet.Name = "Operador"
        Worksheets("Hoja2").Name = "Graficas Operador"
        Worksheets("Graficas Operador").Activate

    End If
    
    Range("A1").Value = "Activo"
    Range("B1").Value = "Suma de Km"
    Range("C1").Value = "Suma de Horas"
    Range("D1").Value = "Suma de Exceso de velocidad 20 Km/H"
    Range("E1").Value = "Máx. de Intensidad (Km/H)"
    Range("F1").Value = "Suma de Duración (Seg)"
    Range("G1").Value = "Suma de Exceso de velocidad 40 Km/H"
    Range("H1").Value = "Máx. de Intensidad (Km/H)"
    Range("I1").Value = "Suma de Duración (Seg)"
    Range("J1").Value = "Suma de Exceso de Velocidad 60 Km/H"
    Range("K1").Value = "Máx. de Intensidad (Km/H)"
    Range("L1").Value = "Suma de Duración (Seg)"
    Range("M1").Value = "Suma de Exceso de velocidad 80 Km/H"
    Range("N1").Value = "Máx. de Intensidad (Km/H)"
    Range("O1").Value = "Suma de Duración (Seg)"
    Range("P1").Value = "Suma de Aceleración Brusca"
    Range("Q1").Value = "Suma de Frenada Brusca"
    Range("R1").Value = "Suma de Conductor NO Identificado"
    Range("S1").Value = "Total Faltas"
    Range("T1").Value = "Total Calificación"

    
End Sub

Private Sub Datos_graficaActivo()
    
    Dim cell_A, cell_B, cell_D, cell_E, cell_F, cell_G, cell_H, cell_I, cell_J, cell_K, cell_L, cell_M, cell_N, cell_O, cell_P, cell_Q, cell_R, cell_S, cell_T, cell_U, cell_V As String
    Dim n2, cell2_A, cell2_B, cell2_C, cell2_D, cell2_E, cell2_F, cell2_G, cell2_H, cell2_I, cell2_J, cell2_K, cell2_L, cell2_M, cell2_N, cell2_O, cell2_P, cell2_Q, cell2_R, cell2_S, cell2_T As String
    Dim c, x, x2, sum_pmedio, promedio, cont_pmedio As Integer
    Dim sum_km, sum_hr, sum_ex20, sum_d1, sum_ex40, sum_d2, sum_ex60, sum_tfaltas, max_int1, max_int2, max_int3, max_int4, Auxmax_int1, Auxmax_int2, Auxmax_int3, Auxmax_int4, sum_totalexc As Double
    Dim flag As Boolean
    
    sum_km = 0
    sum_hr = 0
    sum_ex20 = 0
    sum_d1 = 0
    sum_ex40 = 0
    sum_d2 = 0
    sum_ex60 = 0
    sum_d3 = 0
    sum_ex80 = 0
    sum_d4 = 0
    sum_aceb = 0
    sum_freb = 0
    sum_cond = 0
    sum_tfaltas = 0
    promedio = 0
    sum_pmedio = 0
    cont_pmedio = 0
    sum_totalexc = 0
    
    max_int1 = 0
    max_int2 = 0
    max_int3 = 0
    max_int4 = 0
    
    x = 6
    n = CStr(x)
    cell_A = "A" + n
    cell_B = "B" + n
    cell_D = "D" + n
    cell_E = "E" + n
    cell_F = "F" + n
    cell_G = "G" + n
    cell_H = "H" + n
    cell_I = "I" + n
    cell_J = "J" + n
    cell_K = "K" + n
    cell_L = "L" + n
    cell_M = "M" + n
    cell_N = "N" + n
    cell_O = "O" + n
    cell_P = "P" + n
    cell_Q = "Q" + n
    cell_R = "R" + n
    cell_S = "S" + n
    cell_T = "T" + n
    cell_U = "U" + n
    cell_V = "V" + n

    
    x2 = 2
    n2 = CStr(x2)
    cell2_A = "A" + n2
    cell2_C = "C" + n2
    cell2_B = "B" + n2
    cell2_D = "D" + n2
    cell2_E = "E" + n2
    cell2_F = "F" + n2
    cell2_G = "G" + n2
    cell2_H = "H" + n2
    cell2_I = "I" + n2
    cell2_J = "J" + n2
    cell2_K = "K" + n2
    cell2_L = "L" + n2
    cell2_M = "M" + n2
    cell2_N = "N" + n2
    cell2_O = "O" + n2
    cell2_P = "P" + n2
    cell2_Q = "Q" + n2
    cell2_R = "R" + n2
    cell2_S = "S" + n2
    cell2_T = "T" + n2
    cell2_U = "U" + n2
    cell2_V = "V" + n2

    
    flag = False
    
    Worksheets("Activo").Activate
    

    c = 0

    cantidad_activos = TextBox7.Value
    'MsgBox cantidad_activos
    cantidad_activos = cantidad_activos - 1
    Do While c <= cantidad_activos '11
        
        Do

            If Range(cell_A).Value <> "SUB-TOTAL" Then
                'MsgBox Range(cell_D).Value
                'km = Range(cell_D).Value
                sum_km = Range(cell_D).Value + sum_km
                sum_hr = Range(cell_E).Value + sum_hr
                sum_ex20 = Range(cell_F).Value + sum_ex20
                sum_d1 = Range(cell_H).Value + sum_d1
                sum_ex40 = Range(cell_I).Value + sum_ex40
                sum_d2 = Range(cell_K).Value + sum_d2
                sum_ex60 = Range(cell_L).Value + sum_ex60
                sum_d3 = Range(cell_N).Value + sum_d3
                sum_ex80 = Range(cell_O).Value + sum_ex80
                sum_d4 = Range(cell_Q).Value + sum_d4
                sum_aceb = Range(cell_R).Value + sum_aceb
                sum_freb = Range(cell_S).Value + sum_freb
                sum_cond = Range(cell_T).Value + sum_cond
                sum_tfaltas = Range(cell_U).Value + sum_tfaltas
                
                'sum_pmedio = Range(cell_V).Value + sum_pmedio
                
                
                Auxmax_int1 = Range(cell_G).Value
                
                If Auxmax_int1 > max_int1 Then
                    'MsgBox "Entró" & max_int4
                    max_int1 = Auxmax_int1
                End If
                
                Auxmax_int2 = Range(cell_J).Value
                If Auxmax_int2 > max_int2 Then
                    'MsgBox "Entró" & max_int4
                    max_int2 = Auxmax_int2
                End If
                
                Auxmax_int3 = Range(cell_M).Value
                If Auxmax_int3 > max_int3 Then
                    'MsgBox "Entró" & max_int4
                    max_int3 = Auxmax_int3
                End If
                
                Auxmax_int4 = Range(cell_P).Value
                If Auxmax_int4 > max_int4 Then
                    'MsgBox "Entró" & max_int4
                    max_int4 = Auxmax_int4
                End If
                
                
                'cont_pmedio = cont_pmedio + 1
               
            End If
            
            'MsgBox sum_pmedio
            
            If Range(cell_A).Value = "SUB-TOTAL" Then
                
                
                flag = True
                'MsgBox cell2_B
                Worksheets("Graficas Activo").Range(cell2_A).Value = Range(cell_B).Value
                Worksheets("Graficas Activo").Range(cell2_B).Value = sum_km
                Worksheets("Graficas Activo").Range(cell2_C).Value = sum_hr
                Worksheets("Graficas Activo").Range(cell2_D).Value = sum_ex20
                Worksheets("Graficas Activo").Range(cell2_F).Value = sum_d1
                Worksheets("Graficas Activo").Range(cell2_G).Value = sum_ex40
                Worksheets("Graficas Activo").Range(cell2_I).Value = sum_d2
                Worksheets("Graficas Activo").Range(cell2_J).Value = sum_ex60
                Worksheets("Graficas Activo").Range(cell2_L).Value = sum_d3
                Worksheets("Graficas Activo").Range(cell2_M).Value = sum_ex80
                Worksheets("Graficas Activo").Range(cell2_O).Value = sum_d4
                Worksheets("Graficas Activo").Range(cell2_P).Value = sum_aceb
                Worksheets("Graficas Activo").Range(cell2_Q).Value = sum_freb
                Worksheets("Graficas Activo").Range(cell2_R).Value = sum_cond
                Worksheets("Graficas Activo").Range(cell2_S).Value = sum_tfaltas
                
                'promedio = sum_pmedio / cont_pmedio
                'promedio = CInt(promedio)
                
                sum_totalexc = 100 - ((sum_ex20 + sum_ex40 + sum_ex60 + sum_ex80) * 2)
                
                If sum_totalexc < 0 Then
                    sum_totalexc = 0
                End If
                
                
                Worksheets("Graficas Activo").Range(cell2_T).Value = sum_totalexc 'promedio 'sum_pmedio
                
                Worksheets("Graficas Activo").Range(cell2_E).Value = max_int1
                Worksheets("Graficas Activo").Range(cell2_H).Value = max_int2
                Worksheets("Graficas Activo").Range(cell2_K).Value = max_int3
                Worksheets("Graficas Activo").Range(cell2_N).Value = max_int4
                
                'MsgBox sum_km
                x2 = x2 + 1
                n2 = CStr(x2)
                cell2_A = "A" + n2
                cell2_B = "B" + n2
                cell2_C = "C" + n2
                cell2_D = "D" + n2
                cell2_E = "E" + n2
                cell2_F = "F" + n2
                cell2_G = "G" + n2
                cell2_H = "H" + n2
                cell2_I = "I" + n2
                cell2_J = "J" + n2
                cell2_K = "K" + n2
                cell2_L = "L" + n2
                cell2_M = "M" + n2
                cell2_N = "N" + n2
                cell2_O = "O" + n2
                cell2_P = "P" + n2
                cell2_Q = "Q" + n2
                cell2_R = "R" + n2
                cell2_S = "S" + n2
                cell2_T = "T" + n2
                cell2_U = "U" + n2
                cell2_V = "V" + n2

                
                sum_km = 0
                sum_hr = 0
                sum_ex20 = 0
                sum_d1 = 0
                sum_ex40 = 0
                sum_d2 = 0
                sum_ex60 = 0
                sum_d3 = 0
                sum_ex80 = 0
                sum_d4 = 0
                sum_aceb = 0
                sum_freb = 0
                sum_cond = 0
                sum_tfaltas = 0
                promedio = 0
                sum_pmedio = 0
                cont_pmedio = 0
                sum_totalexc = 0
                max_int1 = 0
                max_int2 = 0
                max_int3 = 0
                max_int4 = 0
                
            End If

            x = x + 1
            n = CStr(x)
            cell_A = "A" + n
            cell_B = "B" + n
            cell_D = "D" + n
            cell_E = "E" + n
            cell_F = "F" + n
            cell_G = "G" + n
            cell_H = "H" + n
            cell_I = "I" + n
            cell_J = "J" + n
            cell_K = "K" + n
            cell_L = "L" + n
            cell_M = "M" + n
            cell_N = "N" + n
            cell_O = "O" + n
            cell_P = "P" + n
            cell_Q = "Q" + n
            cell_R = "R" + n
            cell_S = "S" + n
            cell_T = "T" + n
            cell_U = "U" + n
            cell_V = "V" + n
            
        Loop Until flag
        
    c = c + 1
    flag = False
    Loop
    
    
     
End Sub

Private Sub Datos_graficaOperador()

    Dim cell_A, cell_B, cell_D, cell_E, cell_F, cell_G, cell_H, cell_I, cell_J, cell_K, cell_L, cell_M, cell_N, cell_O, cell_P, cell_Q, cell_R, cell_S, cell_T, cell_U, cell_V, cell_W, cell_X, cell_Y, cell_Z As String
    Dim n2, cell2_A, cell2_B, cell2_C, cell2_D, cell2_E, cell2_F, cell2_G, cell2_H, cell2_I, cell2_J, cell2_K, cell2_L, cell2_M, cell2_N, cell2_O, cell2_P, cell2_Q, cell2_R, cell2_S, cell2_T As String
    Dim c, x, x2, sum_pmedio, promedio, cont_pmedio As Integer
    Dim sum_km, sum_hr, sum_ex20, sum_d1, sum_ex40, sum_d2, sum_ex60, sum_tfaltas, max_int1, max_int2, max_int3, max_int4, Auxmax_int1, Auxmax_int2, Auxmax_int3, Auxmax_int4, sum_totalexc As Double
    Dim flag As Boolean
    
    sum_km = 0
    sum_hr = 0
    sum_ex20 = 0
    sum_d1 = 0
    sum_ex40 = 0
    sum_d2 = 0
    sum_ex60 = 0
    sum_d3 = 0
    sum_ex80 = 0
    sum_d4 = 0
    sum_aceb = 0
    sum_freb = 0
    sum_cond = 0
    sum_tfaltas = 0
    promedio = 0
    sum_pmedio = 0
    cont_pmedio = 0
    sum_totalexc = 0
    
    max_int1 = 0
    max_int2 = 0
    max_int3 = 0
    max_int4 = 0
    
    x = 6
    n = CStr(x)
    cell_A = "A" + n
    cell_B = "B" + n
    cell_D = "D" + n
    cell_E = "E" + n
    cell_F = "F" + n
    cell_G = "G" + n
    cell_H = "H" + n
    cell_I = "I" + n
    cell_J = "J" + n
    cell_K = "K" + n
    cell_L = "L" + n
    cell_M = "M" + n
    cell_N = "N" + n
    cell_O = "O" + n
    cell_P = "P" + n
    cell_Q = "Q" + n
    cell_R = "R" + n
    cell_S = "S" + n
    cell_T = "T" + n
    cell_U = "U" + n
    cell_V = "V" + n
    cell_W = "W" + n
    cell_X = "X" + n
    cell_Y = "Y" + n
    cell_Z = "Z" + n
    
    x2 = 2
    n2 = CStr(x2)
    cell2_A = "A" + n2
    cell2_C = "C" + n2
    cell2_B = "B" + n2
    cell2_D = "D" + n2
    cell2_E = "E" + n2
    cell2_F = "F" + n2
    cell2_G = "G" + n2
    cell2_H = "H" + n2
    cell2_I = "I" + n2
    cell2_J = "J" + n2
    cell2_K = "K" + n2
    cell2_L = "L" + n2
    cell2_M = "M" + n2
    cell2_N = "N" + n2
    cell2_O = "O" + n2
    cell2_P = "P" + n2
    cell2_Q = "Q" + n2
    cell2_R = "R" + n2
    cell2_S = "S" + n2
    cell2_T = "T" + n2
    cell2_U = "U" + n2
    cell2_V = "V" + n2

    
    flag = False
    
    Worksheets("Operador").Activate
    

    c = 0

    cantidad_activos = TextBox7.Value
    'MsgBox cantidad_activos
    cantidad_activos = cantidad_activos - 1
    Do While c <= cantidad_activos '11
        
        Do

            If Range(cell_A).Value <> "SUB-TOTAL" Then
                'MsgBox Range(cell_D).Value
                'km = Range(cell_D).Value
                sum_km = Range(cell_H).Value + sum_km
                sum_hr = Range(cell_I).Value + sum_hr
                sum_ex20 = Range(cell_J).Value + sum_ex20
                sum_d1 = Range(cell_L).Value + sum_d1
                sum_ex40 = Range(cell_M).Value + sum_ex40
                sum_d2 = Range(cell_O).Value + sum_d2
                sum_ex60 = Range(cell_P).Value + sum_ex60
                sum_d3 = Range(cell_R).Value + sum_d3
                sum_ex80 = Range(cell_S).Value + sum_ex80
                sum_d4 = Range(cell_U).Value + sum_d4
                sum_aceb = Range(cell_V).Value + sum_aceb
                sum_freb = Range(cell_W).Value + sum_freb
                sum_cond = Range(cell_X).Value + sum_cond
                sum_tfaltas = Range(cell_U).Value + sum_tfaltas
                
                'sum_pmedio = Range(cell_V).Value + sum_pmedio
                
                
                Auxmax_int1 = Range(cell_K).Value
                
                If Auxmax_int1 > max_int1 Then
                    'MsgBox "Entró" & max_int4
                    max_int1 = Auxmax_int1
                End If
                
                Auxmax_int2 = Range(cell_N).Value
                If Auxmax_int2 > max_int2 Then
                    'MsgBox "Entró" & max_int4
                    max_int2 = Auxmax_int2
                End If
                
                Auxmax_int3 = Range(cell_Q).Value
                If Auxmax_int3 > max_int3 Then
                    'MsgBox "Entró" & max_int4
                    max_int3 = Auxmax_int3
                End If
                
                Auxmax_int4 = Range(cell_T).Value
                If Auxmax_int4 > max_int4 Then
                    'MsgBox "Entró" & max_int4
                    max_int4 = Auxmax_int4
                End If
                
                
                'cont_pmedio = cont_pmedio + 1
               
            End If
            
            'MsgBox sum_pmedio
            
            If Range(cell_A).Value = "SUB-TOTAL" Then
                
                
                flag = True
                'MsgBox cell2_B
                Worksheets("Graficas Operador").Range(cell2_A).Value = Range(cell_B).Value
                Worksheets("Graficas Operador").Range(cell2_B).Value = sum_km
                Worksheets("Graficas Operador").Range(cell2_C).Value = sum_hr
                Worksheets("Graficas Operador").Range(cell2_D).Value = sum_ex20
                Worksheets("Graficas Operador").Range(cell2_F).Value = sum_d1
                Worksheets("Graficas Operador").Range(cell2_G).Value = sum_ex40
                Worksheets("Graficas Operador").Range(cell2_I).Value = sum_d2
                Worksheets("Graficas Operador").Range(cell2_J).Value = sum_ex60
                Worksheets("Graficas Operador").Range(cell2_L).Value = sum_d3
                Worksheets("Graficas Operador").Range(cell2_M).Value = sum_ex80
                Worksheets("Graficas Operador").Range(cell2_O).Value = sum_d4
                Worksheets("Graficas Operador").Range(cell2_P).Value = sum_aceb
                Worksheets("Graficas Operador").Range(cell2_Q).Value = sum_freb
                Worksheets("Graficas Operador").Range(cell2_R).Value = sum_cond
                Worksheets("Graficas Operador").Range(cell2_S).Value = sum_tfaltas
                
                'promedio = sum_pmedio / cont_pmedio
                'promedio = CInt(promedio)
                
                sum_totalexc = 100 - ((sum_ex20 + sum_ex40 + sum_ex60 + sum_ex80) * 2)
                
                If sum_totalexc < 0 Then
                    sum_totalexc = 0
                End If
                
                
                Worksheets("Graficas Operador").Range(cell2_T).Value = sum_totalexc 'promedio 'sum_pmedio
                
                Worksheets("Graficas Operador").Range(cell2_E).Value = max_int1
                Worksheets("Graficas Operador").Range(cell2_H).Value = max_int2
                Worksheets("Graficas Operador").Range(cell2_K).Value = max_int3
                Worksheets("Graficas Operador").Range(cell2_N).Value = max_int4
                
                'MsgBox sum_km
                x2 = x2 + 1
                n2 = CStr(x2)
                cell2_A = "A" + n2
                cell2_B = "B" + n2
                cell2_C = "C" + n2
                cell2_D = "D" + n2
                cell2_E = "E" + n2
                cell2_F = "F" + n2
                cell2_G = "G" + n2
                cell2_H = "H" + n2
                cell2_I = "I" + n2
                cell2_J = "J" + n2
                cell2_K = "K" + n2
                cell2_L = "L" + n2
                cell2_M = "M" + n2
                cell2_N = "N" + n2
                cell2_O = "O" + n2
                cell2_P = "P" + n2
                cell2_Q = "Q" + n2
                cell2_R = "R" + n2
                cell2_S = "S" + n2
                cell2_T = "T" + n2
                cell2_U = "U" + n2
                cell2_V = "V" + n2

                
                sum_km = 0
                sum_hr = 0
                sum_ex20 = 0
                sum_d1 = 0
                sum_ex40 = 0
                sum_d2 = 0
                sum_ex60 = 0
                sum_d3 = 0
                sum_ex80 = 0
                sum_d4 = 0
                sum_aceb = 0
                sum_freb = 0
                sum_cond = 0
                sum_tfaltas = 0
                promedio = 0
                sum_pmedio = 0
                cont_pmedio = 0
                sum_totalexc = 0
                max_int1 = 0
                max_int2 = 0
                max_int3 = 0
                max_int4 = 0
                
            End If

            x = x + 1
            n = CStr(x)
            cell_A = "A" + n
            cell_B = "B" + n
            cell_D = "D" + n
            cell_E = "E" + n
            cell_F = "F" + n
            cell_G = "G" + n
            cell_H = "H" + n
            cell_I = "I" + n
            cell_J = "J" + n
            cell_K = "K" + n
            cell_L = "L" + n
            cell_M = "M" + n
            cell_N = "N" + n
            cell_O = "O" + n
            cell_P = "P" + n
            cell_Q = "Q" + n
            cell_R = "R" + n
            cell_S = "S" + n
            cell_T = "T" + n
            cell_U = "U" + n
            cell_V = "V" + n
            
        Loop Until flag
        
    c = c + 1
    flag = False
    Loop


End Sub


Private Sub diseño_tabla2()

    Dim rango As Integer
    Dim cell, n As String
    
    rango = Contador_de_filas(2) + 2
    n = CStr(rango)
    'Worksheets("Graficas Activo").Activate
    cell = "A1:T" + n
    Range("A1:T1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").EntireRow.AutoFit
    
    
    Range(cell).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    'Range("A1").Select
    
    
    Range("A1:T1").Interior.Color = RGB(209, 229, 254) 'AZUL
    
    Range("A1:T1").Select
    With Selection
         .Font.Bold = True
         .Font.Size = 9
         .Font.Name = "Helvetica"
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True


End Sub

Private Sub tratamiento_tabla2(ByVal cual_informe As String)
    Dim promedio2, rango, contar_filas As Integer
    Dim cell_A, cell_B, cell_C, cell_D, cell_E, cell_F, cell_G, cell_H, cell_I, cell_J, cell_K, cell_L, cell_M, cell_N, cell_O, cell_P, cell_Q, cell_R, cell_S, cell_T, n As String
    
   
    Worksheets(cual_informe).Activate
    rango = Contador_de_filas(2)
    contar_filas = rango + 2
    'MsgBox contar_filas
    
    n = CStr(contar_filas)
    cell_A = "A" + n
    cell_B = "B" + n
    cell_C = "C" + n
    cell_D = "D" + n
    cell_E = "E" + n
    cell_F = "F" + n
    cell_G = "G" + n
    cell_H = "H" + n
    cell_I = "I" + n
    cell_J = "J" + n
    cell_K = "K" + n
    cell_L = "L" + n
    cell_M = "M" + n
    cell_N = "N" + n
    cell_O = "O" + n
    cell_P = "P" + n
    cell_Q = "Q" + n
    cell_R = "R" + n
    cell_S = "S" + n
    cell_T = "T" + n
    
    
    Range(cell_A).Value = "Total"
    rango = rango * (-1)
    n = CStr(rango)
    
    Range(cell_B).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_C).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_D).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_E).Select
    ActiveCell.FormulaR1C1 = "=MAX(R[" + n + "]C:R[-1]C)"
    Range(cell_F).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_G).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_H).Select
    ActiveCell.FormulaR1C1 = "=MAX(R[" + n + "]C:R[-1]C)"
    Range(cell_I).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_J).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_K).Select
    ActiveCell.FormulaR1C1 = "=MAX(R[" + n + "]C:R[-1]C)"
    Range(cell_L).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_M).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_N).Select
    ActiveCell.FormulaR1C1 = "=MAX(R[" + n + "]C:R[-1]C)"
    Range(cell_O).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_P).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_Q).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_R).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    Range(cell_S).Select
    ActiveCell.FormulaR1C1 = "=SUM(R[" + n + "]C:R[-1]C)"
    
    Range(cell_T).Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[" + n + "]C:R[-1]C)"
    
    'promedio2 = Range(cell_T).Value
    'promedio2 = CInt(promedio2)
    'Range(cell_T).Value = promedio2


End Sub

Private Sub graficar(ByVal cual_informe As String)
    
    Dim filas As Integer
    Dim n, cell_A, cell_C, cell, rango, cell_D, cell_F, cell_H, cell_J, cell_L, cell_N, cell_P, cell_Q, cell_R As String

    'GRAFICO 1-------------------------------------------------------------------------------
    
    'filas = Contador_de_filas(6) + 6
    filas = Contador_de_filas(2) + 2
    'MsgBox filas
    n = CStr(filas)
    cell_A = "A" + n
    cell_C = "C" + n
    
    cell_D = "D" + n
    cell_F = "F" + n
    cell_H = "H" + n
    cell_J = "J" + n
    cell_L = "L" + n
    cell_N = "N" + n
    cell_P = "P" + n
    cell_Q = "Q" + n
    cell_R = "R" + n
   
    
    rango = "A1:C1," + cell_A + ":" + cell_C
    Range(rango).Select
    Range(cell_A).Activate
    cell = "$A$" + n + ":$C$" + n
    ActiveSheet.Shapes.AddChart2(286, xl3DColumnClustered).Select
    
    ActiveChart.SetSourceData Source:=Range( _
        "'" + cual_informe + "'!$A$1:$C$1,'" + cual_informe + "'!" + cell)
        
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.ChartTitle.Text = "KILOMETROS vs HORAS"
    ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    
    filas = filas + 1
    n = CStr(filas)
    cell_A = "A" + n
    Range(cell_A).Activate
    
    ActiveSheet.ChartObjects("Gráfico 1").Left = ActiveCell.Left
    ActiveSheet.ChartObjects("Gráfico 1").Top = ActiveCell.Top
    
    
        
    'GRAFICO 2-------------------------------------------------------------------------------
    
    Range("D1,F1,H1,J1,L1,N1,P1,Q1,R1," + cell_D + "," + cell_F + "," + cell_H + "," + cell_J + "," + cell_L + "," + cell_N + "," + cell_P + "," + cell_Q + "," + cell_R).Select
    Range(cell_R).Activate
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveSheet.ChartObjects("Gráfico 2").Activate
    ActiveChart.ChartTitle.Text = "FALTAS"
    ActiveChart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    

    ActiveSheet.ChartObjects("Gráfico 2").Left = ActiveSheet.ChartObjects("Gráfico 1").Width + 1
    ActiveSheet.ChartObjects("Gráfico 2").Top = ActiveSheet.ChartObjects("Gráfico 1").Top


    'MsgBox ActiveCell.Top
    'MsgBox ActiveCell.Left
    
    'MsgBox ActiveCell.Top
    'MsgBox ActiveCell.Left
    'MsgBox ActiveCell.Address(0, 0)
    'MsgBox ActiveCell.Address

    'ActivateSheet.Shapes("Gráfico 2").Range ("A1")
    'ActiveSheet.Shapes("Gráfico 2").IncrementLeft 686.25
    'ActiveSheet.Shapes("Gráfico 2").IncrementTop -245.25
    
    'ActiveChart.SetSourceData Source:=Range("$D$1,$F$1,$H$1,$J$1,$L$1,$N$1,$P$1,$Q$1,$R$1,$D$14,$F$14,$H$14,$J$14,$L$14,$N$14,$Q$14,$R$14")
    'ActiveSheet.ChartObjects("Gráfico 1").Activate
    'ActiveChart.SetSourceData Source:=Range("$D$1,$F$1,$H$1,$J$1,$L$1,$N$1,$P$1,$Q$1,$R$1,$D$14,$F$14,$H$14,$J$14,$L$14,$N$14,$Q$14,$R$14")
    'ActiveChart.ApplyLayout (7)
    
    'Range("$D$1,$F$1,$H$1,$J$1,$L$1,$N$1,$P$1,$Q$1,$R$1,$D$14,$F$14,$H$14,$J$14,$L$14,$N$14,$Q$14,$R$14").Select


End Sub

