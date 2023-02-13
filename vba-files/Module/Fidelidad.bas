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
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Reacci�n y/o fisi�n y/o fusi�n y/o irradiaci�n nuclear, contaminaci�n radioactiva por combustibles nucleares o desechos radiactivos, debidos a su propia combusti�n. "
    Range("F3").Value = "Cr�ditos o pr�stamos concedidos por el asegurado. "
    Range("F4").Value = "Huelga, mot�n y conmoci�n civil."
    Range("F5").Value = "Actos dolosos del empleado cometidos sin fines de lucro."
    Range("F6").Value = "Actos cometidos con la participaci�n del Asegurado, sus accionistas, representantes, propietarios o familiares de cualquiera de los anteriores."
    Range("F7").Value = "La imposibilidad de la empresa asegurada de recuperar los cr�ditos otorgados sin suficiente garant�a o en condiciones irregulares por parte de sus empleados. "
    Range("F8").Value = "Actos de infidelidad cometidos por el empleado, de los cuales tuvo conocimiento el asegurado y no lo inform� al Instituto; ni tom� las acciones necesarias para evitar la consecuci�n del il�cito. "
    Range("F9").Value = "Quiebra e insolvencia de la Empresa Asegurada. "
    Range("F10").Value = "Falta de discreci�n de empleados que causen p�rdidas monetarias al Asegurado."
    Range("F11").Value = "Sanciones pecuniarias o multas a cargo del Asegurado (por ejemplo, multas contractuales) a�n si resultan de actos de infidelidad."
    
    Range("F13").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
