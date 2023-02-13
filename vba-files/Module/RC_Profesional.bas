Attribute VB_Name = "RC_Profesional"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("B1").Value = "RESPONSABILIDAD CIVIL"
    Range("B2").Value = "COBERTURAS"
    Range("B3").Value = "L: RESPONSABILIDAD CIVIL "

    

    Range("C2").Value = "DEDUCIBLES"
    Range("C3").Value = "No contratada"
    
    Range("B5").Value = "Condiciones Particulares"
    Range("B6").Value = "Inserte Condiciones Particulares"
    
    Range("B8").Value = "Condiciones Generales"
    Range("B9").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNo9DOVBOPpUD6AhIA?e=TaLRBo"
    
    Range("B11").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerra, invasi�n, actos de enemigos extranjeros, actividades u operaciones militares."
    Range("F3").Value = "Reacci�n nuclear, irradiaci�n nuclear o contaminaci�n radiactiva por combustibles nucleares o desechos radiactivos."
    Range("F4").Value = "Actos deliberadamente perjudiciales, actos mal intencionados o cometidos con dolo por parte del Asegurado y/o Tomador."
    Range("F5").Value = "Reclamaciones de la que el Asegurado y/o Tomador hubiera tenido conocimiento en el momento de formalizar el contrato."
    Range("F6").Value = "La responsabilidad civil que surja por la p�rdida o da�os resultantes de la explosi�n de una caldera de vapor, u otra clase de recipientes a presi�n concebidos para operar este sistema, que pertenezca al Asegurado, o sea utilizado por �l."
    Range("F7").Value = "Los da�os derivados del indebido ejercicio profesional del Asegurado."
    Range("F8").Value = "Reclamaciones y Demandas provenientes del Exterior."
    Range("F9").Value = "Retraso, p�rdida de mercado u otros da�os consecuenciales."
    Range("F10").Value = "Las obligaciones legalmente imputables al Asegurado bajo la  Legislaci�n de Riesgos del Trabajo."
    Range("F11").Value = "La responsabilidad cubierta mediante contrato de garant�a del fabricante, distribuidor o instalador, o mediante contrato de mantenimiento de los ascensores en uso en el predio asegurado."
    Range("F12").Value = "Da�os ocasionados por profesionales no declarados en las Condiciones Particulares de este seguro."
    Range("F13").Value = "Eventos de la naturaleza."
    Range("F14").Value = "Reclamaciones derivadas de situaciones en que concurra fuerza mayor o derivadas del ejercicio de actividad profesional distinta a la declarada en la solicitud del presente contrato, as� como todas aquellas operaciones ajenas al �mbito estricto de �sta."
    Range("F15").Value = "El empleo, uso o manejo de mercanc�as o productos manufacturados, vendidos, manejados o distribuidos por el Asegurado, cuando exista en ellos una condici�n defectuosa."
    Range("F16").Value = "P�rdidas consecuenciales sufridas por el Asegurado."
    Range("F17").Value = "Insatisfacci�n en la calidad o atributos del servicio profesional prestado."
    Range("F18").Value = "Deficiencias o diferencias del servicio profesional prestado con la promoci�n o publicidad e intereses difusos del mismo."

    Range("B15").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
    'Formato de todo
End Sub
