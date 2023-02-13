Attribute VB_Name = "Incendio_Todo_Riesgo_Colones"

Sub Lafise()
    Range("B1").Value = "Incendio"
    Range("B2").Value = "A: Riesgos No Catastr�ficos"
    Range("B3").Value = "B: Riesgos Catastr�ficos"
    Range("B4").Value = "C: P�rdida de Beneficios Comercial o Industrial"
    Range("B5").Value = "E: P�rdida de Rentas por Contrato de Arrendamiento"
    Range("B6").Value = "D: Gastos Extra"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    
    Range("B8").Value = "Condiciones Particulares"
    Range("B9").Value = "Inserte Condiciones Particulares"
    
    Range("B11").Value = "Condiciones Generales"
    Range("B12").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihPpQFH57NwWMJDBHOw?e=6vhGil"
    
    Range("B14").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "La imposibilidad econ�mica del Tomador y/o Asegurado para hacer frente al gasto de reconstrucci�n, o reparaci�n de la propiedad asegurada."
    Range("F3").Value = "Huelgas, paros, disturbios de car�cter obrero o motines que interrumpan la reconstrucci�n o reparaci�n de la propiedad asegurada o que impidan su uso u ocupaci�n."
    Range("F4").Value = "La aplicaci�n de mandatos o leyes de autoridad competente, salvo lo previsto en la secci�n II de �mbito de Coberturas."
    Range("F5").Value = "Suspensi�n, vencimiento o cancelaci�n de permisos, licencias, contratos de arrendamiento o concesi�n."
    Range("F6").Value = "Saqueo, ya sea durante o despu�s de un siniestro."
    Range("F7").Value = "Propiedad Personal de Visitantes."
    Range("F8").Value = "Hurto de los Bienes Asegurados, excepto cuando ocurran durante un Incendio."
    Range("F9").Value = "Robo o Tentativa de Robo, en los cuales el Tomador y/o Asegurado o sus socios, sean autores o c�mplices."
    Range("F10").Value = "La responsabilidad legal o contractual del fabricante o proveedor de la maquinaria."
    Range("F11").Value = "Da�os o p�rdidas que ocurran por explosi�n de gases de humo en calderas, hornos y/o instalaciones o equipos integrantes."

    
    Range("F14").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
