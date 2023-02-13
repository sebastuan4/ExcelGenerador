Attribute VB_Name = "RC_Productos"
Sub ASSA()
    Range("B1").Value = "RESPONSABILIDAD CIVIL MODALIDAD PRODUCTOS"
    Range("B2").Value = "Coberturas"
    Range("B3").Value = "BÁSICA EN MODALIDAD PRODUCTOS."
    Range("B4").Value = "GENERAL COMPRENSIVA."
    Range("B5").Value = "CRUZADA."
    Range("B6").Value = "DAÑOS POR INCENDIO."
    
    Range("C2").Value = "Deducibles"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"

    
    Range("B9").Value = "Condiciones Particulares"
    Range("B10").Value = "Inserte Condiciones Particulares"
    
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihN4qFUgStKouIKwtEQ?e=yJaczT"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "La falta de limitación de responsabilidad con el distribuidor."
    Range("F3").Value = "Por daños producidos por inobservancia de leyes, reglamentos, disposiciones gubernamentales o por tolerancia de tal inobservancia por personas aseguradas."
    Range("F4").Value = "Por daños que se produzcan en los propios productos suministrados por el Asegurado como consecuencia de un error o defecto de elaboración."
    Range("F5").Value = "Gastos adicionales, por ejemplo retroventa, petición de devolución o reemplazo de los productos ya suministrados."
    Range("F6").Value = "El conocimiento del defecto o de la nocividad de productos suministrados se considerará como culpa grave o dolo."

    Range("F12").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
