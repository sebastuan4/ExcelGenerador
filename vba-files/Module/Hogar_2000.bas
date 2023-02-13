Attribute VB_Name = "Hogar_2000"
Sub ins()
    Range("B1").Value = "HOGAR 2000"
    Range("B2").Value = "A: INCENDIO Y RAYO"
    Range("B3").Value = "B: RIESGOS VARIOS"
    Range("B4").Value = "C: INUNDACIÓN, DESLIZAMIENTO Y VIENTOS "
    Range("B5").Value = "D: CONVULSIONES DE LA NATURALEZA"
    Range("B6").Value = "H: PÉRDIDA DE RENTAS POR CONTRATO DE ARRENDAMIENTO "
    Range("B7").Value = "I: ROTURA DE CRISTALES"
    Range("B8").Value = "R: GASTOS POR ALQUILER"
    Range("B9").Value = "X:  MULTIASISTENCIA HOGAR EXTENDIDA"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    
    Range("B16").Value = "Condiciones Particulares"
    Range("B17").Value = "Inserte Condiciones Particulares"
    
    Range("B19").Value = "Condiciones Generales"
    Range("B20").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNxYZxD_ZSUX9iqfFw?e=R6iLkV"
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerras, terrorismo, invasiones, actos de enemigos extranjeros."
    Range("F3").Value = "Reacción nuclear, irradiación nuclear o contaminación radiactiva "
    Range("F4").Value = "Armas o instrumentos de guerra utilizando fisión o fusión atómica o nuclear u otro como material o fuerza de reacción o radioactiva."
    Range("F5").Value = "Acciones u omisiones del Asegurado, sus empleados o personas actuando en su representación o a quienes se les haya encargado la custodia de los bienes asegurados, que a criterio del instituto produzcan o agraven las pérdidas."
    Range("F6").Value = "Contaminación"
    Range("F7").Value = "Saqueo después de un siniestro. "
    Range("F8").Value = "Las pérdidas consecuenciales, excepto lo previsto en la Cobertura H “Pérdida de Rentas por Contrato de Arrendamiento” y R “Gastos por Alquiler”. "
    Range("F9").Value = "Dolo del Asegurado y/o Tomador"
    Range("F10").Value = "Cuando el uso del inmueble asegurado es ilícito o contrario a la actividad declarada en el contrato póliza. "
    Range("F11").Value = "Daños que se produzcan por colillas de cigarrillo o similares, a menos que produzcan incendio. "
    Range("F12").Value = "Explosión, a menos que produzca incendio y, en este caso, sólo por las pérdidas o daños que dicho incendio ocasione. "
    Range("F13").Value = "Tifones, huracanes, ciclones, erupciones volcánicas, temblores, terremotos, fuegos subterráneos u otras convulsiones de la naturaleza; actos de incendiarios conectados con los acontecimientos anteriores. "
    
    Range("F22").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

