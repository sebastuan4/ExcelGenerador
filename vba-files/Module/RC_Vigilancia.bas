Attribute VB_Name = "RC_Vigilancia"
Sub ins()
    Range("B1").Value = "Responsabilidad Civil Vigilancia"
    Range("B2").Value = "COBERTURAS"
    Range("B3").Value = "L: RESPONSABILIDAD CIVIL."
    
    Range("C2").Value = "DEDUCIBLES"
    Range("C3").Value = "No contratada"
    
    Range("B6").Value = "Condiciones Particulares"
    Range("B7").Value = "Inserte Condiciones Particulares"
    
    Range("B10").Value = "Condiciones Generales"
    Range("B11").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihN4o0YFIUH9HjrPMIQ?e=IkHme4"
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerra, invasión, actos de enemigos extranjeros, actividades u operaciones militares."
    Range("F3").Value = "Reacción nuclear, irradiación nuclear o contaminación radiactiva por combustibles nucleares o desechos radiactivos."
    Range("F4").Value = "Actos deliberadamente perjudiciales, actos mal intencionados o cometidos con dolo por parte del Asegurado y/o Tomador."
    Range("F5").Value = "Reclamaciones de la que el Asegurado y/o Tomador hubiera tenido conocimiento en el momento de formalizar el contrato."
    Range("F6").Value = "Eventos de la naturaleza."
    Range("F7").Value = "Responsabilidad Civil Contractual."
    Range("F8").Value = "Agentes que no se encuentren acreditados."
    Range("F9").Value = "El incumplimiento de la ley de servicios de seguridad 8395."
    Range("F10").Value = "Las obligaciones legalmente imputables al Asegurado bajo la  Legislación de Riesgos del Trabajo."
    Range("F11").Value = "Siniestros en los cuales el vigilante involucrado no se encuentre debidamente incluido en el contrato de la póliza."
    
    Range("F13").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
