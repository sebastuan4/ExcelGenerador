Attribute VB_Name = "estica"
Dim loc_otra_tabla As Integer
Dim loc_fin_tabla As Integer
Sub buscar_fin_tabla()
    loc_fin_tabla = 0
    Do
        loc_fin_tabla = loc_fin_tabla + 1
    Loop Until ActiveCell.Offset(loc_fin_tabla, 0) = ""
End Sub

Sub buscar_otra_tabla()
    loc_otra_tabla = 0
    Do
        loc_otra_tabla = loc_otra_tabla + 1
    Loop Until ActiveCell.Offset(loc_otra_tabla, 0) <> ""
End Sub
    
Sub dar_estetica()
    'Modificando tamaño columnas
    Columns("F:F").Select
    Selection.ColumnWidth = 94.71
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.ColumnWidth = 58.43
    Columns("D:D").Select
    Selection.ColumnWidth = 1.43
    Columns("E:E").Select
    Selection.ColumnWidth = 2.14
    'Rango de inicio
    Range("B1").Activate
    'Buscando fin de coberturas
    Dim tabla_coberturas As Integer
    Call buscar_fin_tabla
    tabla_coberturas = loc_fin_tabla
    
    'Seleccionando el ultimo lugar encontrado para la otra busqueda
    Range("B" & tabla_coberturas + 1).Activate
    
    'Buscando Condicones Particulares
    Dim condiciones_particulares As Integer
    Call buscar_otra_tabla
    condiciones_particulares = loc_otra_tabla
    condiciones_particulares = tabla_coberturas + 1 + condiciones_particulares
    
    'Seleccionando el ultimo lugar encontrado para la otra busqueda
    Range("B" & condiciones_particulares + 2).Activate
    
    'Buscando Condicones Generales
    Dim condiciones_generales As Integer
    Call buscar_otra_tabla
    condiciones_generales = loc_otra_tabla
    condiciones_generales = condiciones_particulares + 2 + condiciones_generales
    
    Range("B" & condiciones_generales + 2).Activate
    
    'Buscando Disclaimer Informacion 1
    Dim disclaimer As Integer
    Call buscar_otra_tabla
    disclaimer = loc_otra_tabla
    disclaimer = condiciones_generales + 2 + disclaimer
    
    
    'Activando el inicio de la tabla de exclusiones
    Range("F1").Activate
    
    'Buscando Fin exclusiones
    Dim exclusiones As Integer
    Call buscar_fin_tabla
    exclusiones = loc_fin_tabla
    
    'Seleccionando el ultimo lugar encontrado para la otra busqueda
    Range("F" & exclusiones + 1).Activate
    
    'Buscando Disclaimer Informacion 2
    Dim disclaimer_2 As Integer
    Call buscar_otra_tabla
    disclaimer_2 = loc_otra_tabla
    disclaimer_2 = exclusiones + disclaimer_2 + 1
    
    'Guardado de los rangos de las tablas
    Dim rango_tabla_coberturas As String
    Dim rango_tabla_condiciones_p As String
    Dim rango_tabla_condiciones_g As String
    Dim rango_disclaimer_1 As String
    Dim rango_tabla_exclusiones As String
    Dim rango_disclaimer_2 As String
    
    rango_tabla_coberturas = "B1:C" & tabla_coberturas
    rango_tabla_condiciones_p = "B" & condiciones_particulares & ":C" & condiciones_particulares + 1
    rango_tabla_condiciones_g = "B" & condiciones_generales & ":C" & condiciones_generales + 1
    rango_disclaimer_1 = "B" & disclaimer & ":C" & disclaimer
    rango_tabla_exclusiones = "F1:F" & exclusiones
    rango_disclaimer_2 = "F" & disclaimer_2
    'Aplicacion de bordes
    Range(rango_tabla_coberturas & "," & rango_tabla_condiciones_p & "," & rango_tabla_condiciones_g & "," & rango_tabla_exclusiones).Select
    Selection.Borders().LineStyle = xlContinuous
    
    'Titulos
    Range("B1,C1,F1,B" & condiciones_particulares & ",C" & condiciones_particulares & ",B" & condiciones_generales & ",C" & condiciones_generales).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .Size = 16
    End With
    
    'Uniendo
     Range("B" & condiciones_particulares & ":C" & condiciones_particulares).Select
     Selection.Merge
     Range("B" & condiciones_particulares + 1 & ":C" & condiciones_particulares + 1).Select
     Selection.Merge
     Range("B" & condiciones_generales & ":C" & condiciones_generales).Select
     Selection.Merge
     Range("B" & condiciones_generales + 1 & ":C" & condiciones_generales + 1).Select
     Selection.Merge
     Range("B" & disclaimer & ":C" & disclaimer).Select
     Selection.Merge
     
    'Bordes Gruesos
    Range(rango_disclaimer_1 & "," & rango_disclaimer_2).Select
    Selection.Borders().LineStyle = xlContinuous
    Selection.Borders().Weight = xlMedium
        
    'Centrando texto y
    Range("B:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    'Formato a las flechas
    ActiveSheet.Shapes.Range(Array("Curved Left Arrow 1")).Select
    Selection.ShapeRange.ScaleWidth 1.5614035088, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.3806228374, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 0.6987951807, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.8636354092, msoFalse, msoScaleFromTopLeft
    
    
End Sub


