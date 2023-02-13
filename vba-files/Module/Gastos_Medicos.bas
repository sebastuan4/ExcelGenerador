Attribute VB_Name = "Gastos_Medicos"
Sub Mapfre()
    'Insertando algo
    Range("B1").Value = "COLECTIVO DE VIDA"
    Range("B2").Value = "A: MUERTE POR CUALQUIER CAUSA"
    Range("B3").Value = "B: MUERTE ACCIDENTAL Y DESMEMBRAMIENTO"
    Range("B4").Value = "C: INCAPACIDAD TOTAL Y/O PERMANENTE"
    Range("B5").Value = "D: GASTOS FUNERARIOS"
    Range("B6").Value = "H: ADELANTO DE SUMA ASEGURADA POR ENFERMEDAD TERMINAL"
    Range("C2").Value = "No tiene"
    Range("C3").Value = "No tiene"
    Range("C4").Value = "No tiene"
    Range("C5").Value = "No tiene"
    Range("C6").Value = "No tiene"
    Range("B9").Value = "Condiciones Particulares"
    Range("B10").Value = "Inserte Condiciones Particulares"
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNo7chwqQEXNwExd9w?e=39nLuI"
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Cobertura Muerte por cualquier causa: Durante los primeros 24 meses de cobertura, en su sano juicio o no, se cause la muerte- suicidio. Si el fallecimiento del Asegurado,  ocurriera durante los primeros 24 meses de cobertura,  siendo la causa de la muerte el Síndrome de Inmunodeficiencia Adquirida (SIDA) y/o el virus de Inmunodeficiencia Adquirida (VIH)."
    Range("F7").Value = "Muerte accidental e Incapacidad Total: Intento de suicidio, Internamientos médicos Ilícitos, Preexistencias  que no hayan podido pasar desapercibidas"
    Range("F10").Value = "Enfermedad terminal  si es a  consecuencia de accidente. Diagnóstico propio o de familiares aunque sean especialistas. Tumores benignos"
    Range("F13").Value = "Todas las coberturas Preexistencias, Epidemias,. Guerra, Energía Atómica o radiación nuclear, Competencias como conductor, Paracidismo, Vuelos en líneas no regulares, Prácticas deportivas submarinas,  Boxeo profesional, escalamiento de montaña, Desemperño en Fuerza pública, Consumo de alcohol, dorogas o medicamentos sin prescripción médica, servicio en bomberos o polícia. terrorismo."
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub PANAMERICAN()
    'Insertando algo
    Range("B1").Value = "Tipo de seguro"
    
    Range("B2").Value = "Gastos Medicos"
    Range("B3").Value = "Vida"
    
    Range("C1").Value = "Deducibles"
    
    Range("C2").Value = "N/A"
    Range("C3").Value = "N/A"

    
    Range("B6").Value = "Condiciones Particulares"
    Range("B7").Value = "Inserte Condiciones Particulares"
    
    Range("B9").Value = "Condiciones Generales"
    Range("B10").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihN8TH2qxZR8FmtK7Ug?e=tbIs4h"
    
    Range("B12").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Homicidio o tentativa de homicidio o por lesiones causadas intencionalmente por una o varias personas."
    Range("F3").Value = "Insurrección o guerra declarada o no, guerra civil, revolución o cualquiera acción atribuible a éstas. Cualquiera otra clase de desorden público o laboral."
    Range("F4").Value = "Cualquier acto de terrorismo, invasión, sedición, bombardeos, usurpación de poder."
    Range("F5").Value = "Enfermedad física o mental, o debido a algún tratamiento médico o quirúrgico o a diagnóstico de éstos."
    Range("F6").Value = "Daños o muerte causada por armas de fuego, armas corto punzantes, artefactos explosivos y/o incendiarios, cualesquiera sean las circunstancias en que ocurran."
    Range("F7").Value = "Daños o muertes por Iatrogenia médica en casos de tratamientos quirúrgicos o médicos los cuales se demuestren negligencia e impericia por parte de médicos tratantes."
    Range("F8").Value = "Energía nuclear (reacciones nucleares, radiación, contaminación)            "
    Range("F9").Value = "Insurrección o guerra declarada o no, guerra civil, revolución o cualquiera acción atribuible a éstas. Cualquiera otra clase de desorden público o laboral."
    Range("F10").Value = "Enfermedad terminal  si es a  consecuencia de accidente. Diagnóstico propio o de familiares aunque sean especialistas. Tumores benignos"
    Range("F11").Value = "Accidentes ocasionados como consecuencia de que el asegurado sufra ataques cardíacos o epilépticos, síncopes; y los accidentes que se produzcan en estado legal de embriaguez, bajo el efecto de las drogas o en estado de sonambulismo o enajenación mental temporal o permanente."
    
    Range("F12").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub ASSA()
    Range("B1").Value = "Gastos Medicos"
    Range("B2").Value = "GASTOS MEDICOS"
            
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "Ver Condiciones Particulares"
    
    Range("B4").Value = "Condiciones Particulares"
    Range("B5").Value = "Inserte Condiciones Particulares"
    
    Range("B6").Value = "Condiciones Generales"
    Range("B7").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOMLmTcvnCHTyzz7Og?e=BiZRGi"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Algún tratamiento o servicio que no esté especificado dentro de los beneficios del Plan."
    Range("F3").Value = "Padecimientos preexistentes no declarados en la solicitud de seguros."
    Range("F4").Value = "Cualquier servicio o suministro que no sea, a juicio de la Compañía, médicamente necesario para el diagnóstico y/o tratamiento de cualquier enfermedad o lesión accidental."
    Range("F5").Value = "Cualquier lesión o enfermedad en que el Asegurado participe por culpa de él mismo. Lesiones que se produzcan a consecuencia de delitos intencionales de los que sea responsable y/o sea participante el Asegurado."
    Range("F6").Value = "Accidentes sufridos en viajes aéreos salvo que el Asegurado Principal o Familiar Asegurado se encuentre viajando como pasajero."
    Range("F7").Value = "Tratamientos dentales, curas u operaciones odontológicas, que no sean a consecuencia de un accidente sufrido dentro de la vigencia de la póliza, salvo los especificados en las Condiciones Particulares de la póliza."
    Range("F8").Value = "Dolo y/o fraude del Asegurado y/o Tomador."
    Range("F9").Value = "Curas de reposo o descanso, controles periódicos o exámenes generales o rutinarios, vacunaciones, certificaciones médicas, así como cualquier otro examen que no haya sido previamente autorizado por La Compañía."
    Range("F10").Value = "Abortos y legrados uterinos punibles."
    Range("F11").Value = "Enfermedades de transmisión sexual, excepto el Síndrome de Inmunodeficiencia Adquirida (SIDA)."
    Range("F12").Value = "En ningún caso se cubrirán la renta o compra aparatos auditivos."
    Range("F13").Value = "Cualquier gasto realizado fuera de la vigencia de la póliza."
    
    Range("F15").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub





