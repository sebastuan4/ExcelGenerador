Attribute VB_Name = "RT"
Sub ins()
    Range("B1").Value = "Riesgos del trabajo"
    Range("B2").Value = "Accidentes de trabajo: Es el accidente que le ocurra al trabajador, con ocasión o por consecuencia del trabajo que desempeña, durante el tiempo que permanece bajo la dirección y dependencia del patrono o sus representantes en forma subordinada y remunerada."
    Range("B3").Value = "Enfermedad de trabajo: Es todo estado patológico que resulte de la acción continuada de una causa, que tiene su origen o motivo en el propio trabajo o en el medio y condiciones en que el trabajador labora, y debe establecerse que éstos han sido la causa de la enfermedad."
    
    Range("C1").Value = "Deducibles"
    Range("C2").Value = "N/A"
    Range("C3").Value = "N/A"

    Range("B6").Value = "CONDICIONES PARTICULARES Y ESTUDIO DE COSTOS."
    Range("B7").Value = "Inserte Link"
    
    Range("B9").Value = "Condiciones Pariculares"
    Range("B10").Value = "Inserte Link"
    
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihNpLejRueCUA6OUtqg?e=zoofiD"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Consumo de drogas, alcohol o similares"

    Range("F6").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
