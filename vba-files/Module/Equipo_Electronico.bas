Attribute VB_Name = "Equipo_Electronico"
Sub ins()
    'Coberturas Particulares
    Range("B1").Value = "MULTIRIESGO COBERTURAS"
    Range("B2").Value = "A: DA�O DIRECTO EQUIPO ELECTR�NICO"
    Range("B3").Value = "B: ROBO"
    Range("B4").Value = "E: EQUIPO M�VIL Y/O PORT�TIL"
    Range("B5").Value = "C: EVENTOS DE LA NATURALEZA"
    Range("B6").Value = "D: OTROS RIESGOS"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    
    Range("B8").Value = "Condiciones Particulares"
    Range("B9").Value = "Inserte Condiciones Particulares"
    
    Range("B10").Value = "Condiciones Generales"
    Range("B11").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOIK7aulN0gyIrmZMg?e=V1bPgA"
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El efecto de virus inform�tico."
    Range("F3").Value = "Hurto."
    Range("F4").Value = "Infidelidad (incluidos actos dolosos, tales como: hurto, robo, estafa o pillaje) de parte de los empleados del Asegurado causados directamente o en complicidad con otros."
    Range("F5").Value = "El funcionamiento continuo, desgaste, cavitaci�n, erosi�n, corrosi�n, o incrustaciones del equipo asegurado."
    Range("F6").Value = "Faltantes que se descubran al efectuar inventarios f�sicos o revisiones de control."
    Range("F7").Value = "La exposici�n continua a la ca�da de arena o ceniza volc�nica, cuando el Asegurado pueda ejercer control para minimizar o evitar tales p�rdidas."
    Range("F8").Value = "El aterrizaje de cabezas lectoras, que produzca da�os a discos duros."
    
    Range("F13").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
