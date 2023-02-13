Attribute VB_Name = "Incendio_Comercial_Colones"
Sub ins()
    Range("B1").Value = "INCENDIO COMERCIAL COBERTURAS"
    Range("B2").Value = "A: INCENDIO CASUAL Y RAYO"
    Range("B3").Value = "B: RIESGOS VARIOS"
    Range("B4").Value = "C: TODO RIESGO INUNDACI�N, DESLIZAMIENTO Y VIENTOS"
    Range("B5").Value = "D: TODO RIESGO CONVULSIONES DE LA NATURALEZA"
    Range("B6").Value = "E: DA�O DIRECTO A LA MERCANC�A (COBERTURA ADICIONAL �NICAMENTE PARA ALMACENES DE DEP�SITO FISCAL Y/O GENERAL"
    Range("B7").Value = "F: P�RDIDA DE BENEFICIOS"
    Range("B8").Value = "G: LLUVIA Y DERRAME"
    Range("B9").Value = "H: P�RDIDA DE RENTA POR CONTRATO DE ARRENDAMIENTO"
    Range("B10").Value = "I: ROTURA DE CRISTALES"
    Range("B11").Value = "Q: GASTOS EXTRA"
    Range("B12").Value = "R: ROBO O TENTATIVA DE ROBO"
    Range("B13").Value = "X: MULTIASISTENCIA COMERCIAL (PLAN TOTAL PLUS)"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    Range("C10").Value = "No contratada"
    Range("C11").Value = "No contratada"
    Range("C12").Value = "No contratada"
    Range("C13").Value = "No contratada"
    
    Range("B16").Value = "Condiciones Particulares"
    Range("B17").Value = "Inserte Condiciones Particulares"
    
    Range("B19").Value = "Condiciones Generales"
    Range("B20").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNxYZxD_ZSUX9iqfFw?e=R6iLkV"
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerras, terrorismo, invasiones, actos de enemigos extranjeros."
    Range("F3").Value = "Reacci�n nuclear, irradiaci�n nuclear o contaminaci�n radiactiva "
    Range("F4").Value = "Armas o instrumentos de guerra utilizando fisi�n o fusi�n at�mica o nuclear u otro como material o fuerza de reacci�n o radioactiva."
    Range("F5").Value = "Acciones u omisiones del Asegurado, sus empleados o personas actuando en su representaci�n o a quienes se les haya encargado la custodia de los bienes asegurados, que a criterio del instituto produzcan o agraven las p�rdidas."
    Range("F6").Value = "P�rdidas o da�os de la propiedad asegurada por fermentaci�n, vicio propio o combusti�n espont�nea."
    Range("F7").Value = "Saqueo, excepto si el siniestro ocurrido es a consecuencia de un evento amparado en la p�liza."
    Range("F8").Value = "P�rdidas directas que tengan su origen en errores de dise�o o defectos constructivos."
    Range("F9").Value = "Toda p�rdida consecuencial."
    Range("F10").Value = "P�rdidas que se originen por cumplimiento de leyes, ordenanzas o reglamentos."
    Range("F11").Value = "En relaci�n con la partida de mercanc�as, en la protecci�n de localizaci�n m�ltiple, se excluye el riesgo de transporte entre bodegas."
    Range("F12").Value = "Los da�os sufridos por los objetos asegurados que se encuentren fuera de los predios asegurados."
    Range("F13").Value = "Dolo del Asegurado y/o Tomador."
    
    Range("F22").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

