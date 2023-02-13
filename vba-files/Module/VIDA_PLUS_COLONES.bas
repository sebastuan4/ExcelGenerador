Attribute VB_Name = "VIDA_PLUS_COLONES"
Sub ins()
    Range("B1").Value = "VIDA"
    Range("B2").Value = "Muerte accidental o no accidental del Asegurado"
    Range("B3").Value = "A:Doble indemnización por Muerte accidental, desmembramiento o pérdida de la vista por causa accidental"
    Range("B4").Value = "B:Coberturas de exoneración de pago de primas y renta en caso de Incapacidad total y permanente del Asegurado y de Cobertura de Pago adicional de la suma asegurada de la cobertura básica en caso de incapacidad total y permanente pagadera en una cuota. "
    Range("B5").Value = "C:Cobertura de adelanto de la mitad de la suma asegurada de la cobertura básica (AMSA). "
    Range("B6").Value = "D:Cobertura de Seguro Temporal"
    Range("B7").Value = "E:Cobertura de Muerte Accidental y no accidental para “Otro Asegurado”. "
    Range("B8").Value = "F:Cobertura de Muerte Accidental y no accidental para Hijos. "
    Range("B9").Value = "G: Cobertura de Indemnización para gastos funerarios"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    
    Range("B11").Value = "Condiciones Particulares"
    Range("B12").Value = "Inserte Condiciones Particulares"
    
    Range("B13").Value = "Condiciones Generales"
    Range("B14").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOY8OCtOkg5NZ4a-3w?e=iufAQI"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Ver condiciones particulares"
    
    Range("F16").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
