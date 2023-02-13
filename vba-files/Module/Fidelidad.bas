Attribute VB_Name = "Fidelidad"
Sub ins()
    Range("B1").Value = "FIDELIDAD POSICIONES COLONES"
    Range("B2").Value = "A: Fidelidad de Posiciones"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    
    Range("B4").Value = "Condiciones Particulares"
    Range("B5").Value = "Inserte Condiciones Particulares"
    
    Range("B7").Value = "Condiciones Generales"
    Range("B8").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihO879mKi-S2oWhesNg?e=XfBhSx"
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Reacción y/o fisión y/o fusión y/o irradiación nuclear, contaminación radioactiva por combustibles nucleares o desechos radiactivos, debidos a su propia combustión. "
    Range("F3").Value = "Créditos o préstamos concedidos por el asegurado. "
    Range("F4").Value = "Huelga, motín y conmoción civil."
    Range("F5").Value = "Actos dolosos del empleado cometidos sin fines de lucro."
    Range("F6").Value = "Actos cometidos con la participación del Asegurado, sus accionistas, representantes, propietarios o familiares de cualquiera de los anteriores."
    Range("F7").Value = "La imposibilidad de la empresa asegurada de recuperar los créditos otorgados sin suficiente garantía o en condiciones irregulares por parte de sus empleados. "
    Range("F8").Value = "Actos de infidelidad cometidos por el empleado, de los cuales tuvo conocimiento el asegurado y no lo informó al Instituto; ni tomó las acciones necesarias para evitar la consecución del ilícito. "
    Range("F9").Value = "Quiebra e insolvencia de la Empresa Asegurada. "
    Range("F10").Value = "Falta de discreción de empleados que causen pérdidas monetarias al Asegurado."
    Range("F11").Value = "Sanciones pecuniarias o multas a cargo del Asegurado (por ejemplo, multas contractuales) aún si resultan de actos de infidelidad."
    
    Range("F13").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
