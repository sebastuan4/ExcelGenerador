Attribute VB_Name = "Responsabilidad_Umbrella"
Sub OCEANICA()
    Range("B1").Value = "Responsabilidad Umbrella"
    Range("B2").Value = "A (BÁSICA): RESPONSABILIDAD CIVIL EXTRACONTRACTUAL UMBRELLA"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    
    Range("B12").Value = "Condiciones Particulares"
    Range("B13").Value = "Inserte Condiciones Particulares"
    
    Range("B15").Value = "Condiciones Generales"
    Range("B16").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihP1GpnajnLvVNwFRPg?e=dHrqjx"
    
    Range("B19").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Salvo que exista una póliza básica que brinde las siguientes coberturas y que haya sido aceptada por Oceánica conforme conste en las Condiciones Particulares, se excluye expresamente:"
    Range("F3").Value = "Responsabilidad Civil profesional"
    Range("F4").Value = "Responsabilidad civil Directores y Ejecutivos"
    Range("F5").Value = "Responsabilidad civil por contaminación "
    Range("F6").Value = "Responsabilidad civil operadores portuarios y aeroportuarios."
    Range("F7").Value = "Responsabilidad civil productos."
    Range("F8").Value = "Responsabilidad Civil patronal"
    Range("F9").Value = "Responsabilidad Penal"
    Range("F10").Value = "Responsabilidad civil contractual"
    Range("F11").Value = "Responsabilidad Civil por polución y/o contaminación gradual o accidental"
    Range("F12").Value = "Responsabilidad civil por la explosión de calderas"
    Range("F13").Value = "Multas de cualquier tipo o fianzas."
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
