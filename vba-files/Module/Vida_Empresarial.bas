Attribute VB_Name = "Vida_Empresarial"
Sub ASSA()
    Range("B1").Value = "Vida Colectivo"
    Range("B2").Value = "Vida"
            
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "Ver Condiciones Particulares"
    
    Range("B4").Value = "Condiciones Particulares"
    Range("B5").Value = "Inserte Condiciones Particulares"
    
    Range("B7").Value = "Condiciones Generales"
    Range("B8").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOMLmTcvnCHTyzz7Og?e=BiZRGi"
    
    Range("B10").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Suicidio o intento de suicidio, estando o no el Asegurado en uso de sus facultades mentales."
    Range("F3").Value = "Lesiones causadas intencionalmente por una o varias personas o por el propio asegurado."
    Range("F4").Value = "Enfermedad corporal o mental; tratamiento médico quirúrgico, salvo si este es a consecuencia de un accidente."
    Range("F5").Value = "La acción de drogas, alcohol, veneno, gas o vapores tomados, administrados, absorbido o inhalados voluntaria o accidentalmente o de alguna otra forma, y todo acontecimiento que se derive del estado de endrogamiento o de embriaguez del Asegurado."
    Range("F6").Value = "Fenómenos de la naturaleza de carácter catastrófico, tales como huracanes, ciclones, tornados, vendavales, deslizamientos de tierra, erupciones volcánicas, terremotos, maremotos, inundaciones y similares."
    Range("F7").Value = "Tratamientos dentales, curas u operaciones odontológicas, que no sean a consecuencia de un accidente sufrido dentro de la vigencia de la póliza, salvo los especificados en las Condiciones Particulares de la póliza."
    Range("F8").Value = "Lesiones sufridas por el Asegurado mientras participa en la comisión o intento de comisión de asalto, asesinato, atentado, o cualquier otro delito."
    Range("F9").Value = "Tomaínas o infección bacteriana (excepto la infección piogénica), cuando ésta se presenta con y por una cortadura o herida recibida por un accidente."
    
    Range("F10").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
