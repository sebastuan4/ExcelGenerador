Attribute VB_Name = "Carga_Internacional_Dolares"
Sub OCEANICA()
    Range("B1").Value = "Carga Internacional"
    Range("B2").Value = "A: TODO RIESGO"
    Range("B3").Value = "C: RIESGO NOMBRADO"
    Range("B4").Value = "D: CLÁUSULA A DEL INSTITUTO DE LONDRES PARA PRODUCTOS PERECEDEROS Y/O REFRIGERADOS"
    Range("B5").Value = "E: CLÁUSULA C DEL INSTITUTO DE LONDRES PARA PRODUCTOS PERECEDEROS Y/O REFRIGERADOS"
    Range("B6").Value = "F: HUELGA"
    Range("B7").Value = "G: GUERRA"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    
    Range("B9").Value = "Condiciones Particulares"
    Range("B10").Value = "Inserte Condiciones Particulares"
    
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOphw_ii6x2Mvr50Qw?e=UWV8ll"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Violación a cualquier Ley, disposición o reglamento de alguna autoridad constituida, sea internacional, nacional, federal, de Estado, municipal o local."
    Range("F3").Value = "Pérdida, daño o gasto ocasionado por vicio inherente o las propias características de los bienes asegurados."
    Range("F4").Value = "Pérdida, daño o gasto atribuible a actos malintencionados del asegurado."
    Range("F5").Value = "Cierre total o parcial del servicio, falta de ocupación o suspensión de las actividades por un período mayor de un mes, de las edificaciones aseguradas o que contengan los bienes asegurados."
    Range("F6").Value = "Derrame normal u ordinario, pérdida normal de peso, volumen o masa o características por merma y uso o desgaste normal de los bienes asegurados durante el transporte de las mercancía; cuando la merma sea igual o mayor a un 3%"
    Range("F7").Value = "Responsabilidad Civil derivada por contaminación de la carga transportada."
    Range("F8").Value = "Actos de contrabando o comercio ilegal."
    Range("F9").Value = "Polución."
    Range("F10").Value = "Sabotaje."
    Range("F11").Value = "Infidelidad de los empleados del asegurado o de la empresa de transporte"
    Range("F12").Value = "Actos terroristas o personas que actúen por motivaciones políticas;"
    Range("F13").Value = "Pérdida o daños a bienes o personas, resultantes de intentos de intimidación o coerción a un gobierno, población civil, en fomento, avance o promoción de objetivos políticos, sociales o religiosos."
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

