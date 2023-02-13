Attribute VB_Name = "Carga_Dolares"
Sub ASSA()
    Range("B1").Value = "Coberturas Carga"
    Range("B2").Value = "A COBERTURA DE RIESGOS DE LAS CLÁUSULAS DEL INSTITUTO DE  LONDRES PARA CARGAMENTOS “A” (1/1/82), No.252 y 259: "
    Range("B3").Value = "B COBERTURA DE RIESGOS DE LAS CLÁUSULAS DEL INSTITUTO DE  LONDRES PARA CARGAMENTOS “B”  (1/1/82), No.253: "

    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    
    Range("B9").Value = "Condiciones Particulares"
    Range("B10").Value = "Inserte Condiciones Particulares"
    
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihPsuZzf5YURiJQU-vQ?e=i21tzn"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Pérdida, daño o gasto atribuible a falta voluntaria del Asegurado. "
    Range("F3").Value = "Merma normal, pérdida normal de peso o volumen, o uso y desgaste normal del interés asegurado. "
    Range("F4").Value = "Pérdida, daño o gasto causado por vicio propio o por la naturaleza del interés asegurado. "
    Range("F5").Value = "Pérdida, daño o gasto causado próximamente por retraso, aún cuando el retraso sea causado por un riesgo asegurado (excepto gastos pagaderos bajo la Cobertura Básica, Cláusula de Avería General). "
    Range("F6").Value = "Pérdida, daño o gasto que se derive de la insolvencia o incumplimiento financiero penable de los propietarios, administradores, fletadores u operadores del buque. "
    Range("F7").Value = "Pérdida, daño o gasto que se origine del uso de cualquier arma de guerra en la cual se emplee fisión atómica o nuclear y/o fusión u otra reacción similar o fuerza o materia radioactiva. "
    Range("F8").Value = "Huelgas, cierre de fábricas, disturbios laborales, motines o tumultos populares "
    Range("F9").Value = "Minas derrelictas, torpedos, bombas u otras armas de guerra derrelictas. "
    Range("F10").Value = "La falta de condiciones de navegabilidad del buque o embarcación. "
    Range("F11").Value = "Captura, embargo, arresto, restricción o detención (excepto piratería) ni de sus consecuencias o de cualquier intento para ello. "
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
