Attribute VB_Name = "Vida_Individual"
Sub ins()
    Range("B1").Value = "MULTIRIESGO COBERTURAS"
    Range("B2").Value = "A: COBERTURA DE GASTOS M�DICOS."
    Range("B3").Value = "B: COBERTURA DE ASISTENCIA AL VIAJERO."
    Range("B4").Value = "C: COBERTURA DE CHEQUEOS."
    Range("B5").Value = "D: COBERTURA POR FALLECIMIENTO."
    Range("B6").Value = "E: COBERTURA DENTAL POR ACCIDENTE Y/O EMERGENCIA."
    Range("B7").Value = "A Extra: COBERTURA PARA ENFERMEDADES Y ACCIDENTES GRAVES."
    Range("B8").Value = "B Extra: COBERTURA ADICIONAL DE C�NCER."
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    
    Range("B10").Value = "Condiciones Particulares"
    Range("B11").Value = "Inserte Condiciones Particulares"
    
    Range("B13").Value = "Condiciones Generales"
    Range("B14").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihONDEQOT9rhLUawKew?e=dMzFn1"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Accidentes provocados intencionalmente por el Asegurado o en los que no medie la acci�n repentina de un agente externo."
    Range("F3").Value = "El accidente o enfermedad sufrido por el Asegurado como consecuencia de la comisi�n o tentativa de delito doloso en que el mismo sea el sujeto activo."
    Range("F4").Value = "Accidentes ocurridos al Asegurado, con o sin intenci�n, cuando �ste �ltimo se encuentre bajo el efecto del alcohol, drogas o estupefacientes, no prescritos por un m�dico u odont�logo."
    Range("F5").Value = "Accidentes ocurridos al Asegurado mientras conduzca un veh�culo y no cuente con la licencia habilitante (independientemente si se encontrase en la v�a p�blica o no)."
    Range("F6").Value = "Toda condici�n preexistente, excepto lo contemplado en la Cobertura de Gastos M�dicos Sujetos a Subl�mite punto Enfermedades Cong�nitas del reci�n nacido."
    Range("F7").Value = "Tratamientos experimentales. "
    Range("F8").Value = "Toda aquella enfermedad mental no tratada por m�dico con especialidad en psiquiatr�a."
    Range("F9").Value = "Bulimia, anorexia nerviosa, fatiga y estr�s."
    Range("F10").Value = "Enfermedades, condiciones o padecimientos, que se originen como consecuencia del consumo de alcohol, tabaco o uso de drogas il�citas."
    Range("F11").Value = "M�todos anticonceptivos quir�rgicos y no quir�rgicos."
    Range("F12").Value = "Gastos odontol�gicos que no sean producto de una emergencia m�dica."
    Range("F13").Value = " Enfermedades de transmisi�n sexual (ven�reas)."
    
    Range("F16").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
