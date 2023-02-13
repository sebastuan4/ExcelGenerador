Attribute VB_Name = "Medico_Colectivo"
Sub ins()
    Range("B1").Value = "AUTOM�VILES"
    Range("B2").Value = "Vida"
    Range("B3").Value = "Gastos Medicos"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    
    Range("B21").Value = "Condiciones Particulares"
    Range("B22").Value = "Inserte Condiciones Particulares"
    
    Range("B24").Value = "Condiciones Generales"
    Range("B25").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihPAJp_UpYPhLCZgyjQ?e=lB8LCt"
    
    Range("B27").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Accidentes provocados intencionalmente por el Asegurado o en los que no existi� la acci�n repentina de un agente externo."
    Range("F3").Value = "Accidentes ocurridos al Asegurado, con o sin intenci�n, cuando este �ltimo se encuentre bajo el efecto del alcohol, drogas o estupefacientes."
    Range("F4").Value = "Accidentes donde el Asegurado conduzca un veh�culo y no cuente con la licencia habilitante (independientemente si se encontrase en la v�a p�blica o no)."
    Range("F5").Value = "Accidentes a pilotos o miembros de tripulaci�n de aeronaves mientras se encuentre desempe�ando sus funciones laborales."
    Range("F6").Value = "El accidente o enfermedad sufrido por el Asegurado como consecuencia de la comisi�n o tentativa de delito doloso en que el mismo sea el sujeto activo."
    Range("F7").Value = "Guerra internacional declarada o no, guerra civil, invasi�n, terrorismo,insurrecci�n, participaci�n activa en alteraciones del orden p�blico."
    Range("F8").Value = "Sea utilizado en competencias o en pruebas de seguridad."
    Range("F9").Value = "Todo tratamiento no prescrito por un m�dico u odont�logo, o por uno que no se encuentre activo o habilitado en el Colegio Profesional respectivo."
    Range("F10").Value = "Tratamientos experimentales."
    Range("F11").Value = "Bulimia, anorexia nerviosa, fatiga y estr�s."
    Range("F12").Value = "M�todos anticonceptivos quir�rgicos y no quir�rgicos."
    
    Range("F21").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
