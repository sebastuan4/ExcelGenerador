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
    
    Range("B11").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerra, invasión, actos de enemigos extranjeros, actividades u operaciones militares."
    Range("F3").Value = "Reacción nuclear, irradiación nuclear o contaminación radiactiva por combustibles nucleares o desechos radiactivos."
    Range("F4").Value = "Actos deliberadamente perjudiciales, actos mal intencionados o cometidos con dolo por parte del Asegurado y/o Tomador."
    Range("F5").Value = "Reclamaciones de la que el Asegurado y/o Tomador hubiera tenido conocimiento en el momento de formalizar el contrato."
    Range("F6").Value = "La responsabilidad civil que surja por la pérdida o daños resultantes de la explosión de una caldera de vapor, u otra clase de recipientes a presión concebidos para operar este sistema, que pertenezca al Asegurado, o sea utilizado por él."
    Range("F7").Value = "Los daños derivados del indebido ejercicio profesional del Asegurado."
    Range("F8").Value = "Reclamaciones y Demandas provenientes del Exterior."
    Range("F9").Value = "Retraso, pérdida de mercado u otros daños consecuenciales."
    Range("F10").Value = "Las obligaciones legalmente imputables al Asegurado bajo la  Legislación de Riesgos del Trabajo."
    Range("F11").Value = "La responsabilidad cubierta mediante contrato de garantía del fabricante, distribuidor o instalador, o mediante contrato de mantenimiento de los ascensores en uso en el predio asegurado."
    Range("F12").Value = "Daños ocasionados por profesionales no declarados en las Condiciones Particulares de este seguro."
    Range("F13").Value = "Eventos de la naturaleza."
    Range("F14").Value = "Reclamaciones derivadas de situaciones en que concurra fuerza mayor o derivadas del ejercicio de actividad profesional distinta a la declarada en la solicitud del presente contrato, así como todas aquellas operaciones ajenas al ámbito estricto de ésta."
    Range("F15").Value = "El empleo, uso o manejo de mercancías o productos manufacturados, vendidos, manejados o distribuidos por el Asegurado, cuando exista en ellos una condición defectuosa."
    Range("F16").Value = "Pérdidas consecuenciales sufridas por el Asegurado."
    Range("F17").Value = "Insatisfacción en la calidad o atributos del servicio profesional prestado."
    Range("F18").Value = "Deficiencias o diferencias del servicio profesional prestado con la promoción o publicidad e intereses difusos del mismo."

    Range("B15").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
    'Formato de todo
End Sub
