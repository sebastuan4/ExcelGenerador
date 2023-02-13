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
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Cobertura Muerte por cualquier causa: Durante los primeros 24 meses de cobertura, en su sano juicio o no, se cause la muerte- suicidio. Si el fallecimiento del Asegurado,  ocurriera durante los primeros 24 meses de cobertura,  siendo la causa de la muerte el S�ndrome de Inmunodeficiencia Adquirida (SIDA) y/o el virus de Inmunodeficiencia Adquirida (VIH)."
    Range("F7").Value = "Muerte accidental e Incapacidad Total: Intento de suicidio, Internamientos m�dicos Il�citos, Preexistencias  que no hayan podido pasar desapercibidas"
    Range("F10").Value = "Enfermedad terminal  si es a  consecuencia de accidente. Diagn�stico propio o de familiares aunque sean especialistas. Tumores benignos"
    Range("F13").Value = "Todas las coberturas Preexistencias, Epidemias,. Guerra, Energ�a At�mica o radiaci�n nuclear, Competencias como conductor, Paracidismo, Vuelos en l�neas no regulares, Pr�cticas deportivas submarinas,  Boxeo profesional, escalamiento de monta�a, Desemper�o en Fuerza p�blica, Consumo de alcohol, dorogas o medicamentos sin prescripci�n m�dica, servicio en bomberos o pol�cia. terrorismo."
    Range("F18").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
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
    
    Range("B12").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Homicidio o tentativa de homicidio o por lesiones causadas intencionalmente por una o varias personas."
    Range("F3").Value = "Insurrecci�n o guerra declarada o no, guerra civil, revoluci�n o cualquiera acci�n atribuible a �stas. Cualquiera otra clase de desorden p�blico o laboral."
    Range("F4").Value = "Cualquier acto de terrorismo, invasi�n, sedici�n, bombardeos, usurpaci�n de poder."
    Range("F5").Value = "Enfermedad f�sica o mental, o debido a alg�n tratamiento m�dico o quir�rgico o a diagn�stico de �stos."
    Range("F6").Value = "Da�os o muerte causada por armas de fuego, armas corto punzantes, artefactos explosivos y/o incendiarios, cualesquiera sean las circunstancias en que ocurran."
    Range("F7").Value = "Da�os o muertes por Iatrogenia m�dica en casos de tratamientos quir�rgicos o m�dicos los cuales se demuestren negligencia e impericia por parte de m�dicos tratantes."
    Range("F8").Value = "Energ�a nuclear (reacciones nucleares, radiaci�n, contaminaci�n)            "
    Range("F9").Value = "Insurrecci�n o guerra declarada o no, guerra civil, revoluci�n o cualquiera acci�n atribuible a �stas. Cualquiera otra clase de desorden p�blico o laboral."
    Range("F10").Value = "Enfermedad terminal  si es a  consecuencia de accidente. Diagn�stico propio o de familiares aunque sean especialistas. Tumores benignos"
    Range("F11").Value = "Accidentes ocasionados como consecuencia de que el asegurado sufra ataques card�acos o epil�pticos, s�ncopes; y los accidentes que se produzcan en estado legal de embriaguez, bajo el efecto de las drogas o en estado de sonambulismo o enajenaci�n mental temporal o permanente."
    
    Range("F12").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
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
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Alg�n tratamiento o servicio que no est� especificado dentro de los beneficios del Plan."
    Range("F3").Value = "Padecimientos preexistentes no declarados en la solicitud de seguros."
    Range("F4").Value = "Cualquier servicio o suministro que no sea, a juicio de la Compa��a, m�dicamente necesario para el diagn�stico y/o tratamiento de cualquier enfermedad o lesi�n accidental."
    Range("F5").Value = "Cualquier lesi�n o enfermedad en que el Asegurado participe por culpa de �l mismo. Lesiones que se produzcan a consecuencia de delitos intencionales de los que sea responsable y/o sea participante el Asegurado."
    Range("F6").Value = "Accidentes sufridos en viajes a�reos salvo que el Asegurado Principal o Familiar Asegurado se encuentre viajando como pasajero."
    Range("F7").Value = "Tratamientos dentales, curas u operaciones odontol�gicas, que no sean a consecuencia de un accidente sufrido dentro de la vigencia de la p�liza, salvo los especificados en las Condiciones Particulares de la p�liza."
    Range("F8").Value = "Dolo y/o fraude del Asegurado y/o Tomador."
    Range("F9").Value = "Curas de reposo o descanso, controles peri�dicos o ex�menes generales o rutinarios, vacunaciones, certificaciones m�dicas, as� como cualquier otro examen que no haya sido previamente autorizado por La Compa��a."
    Range("F10").Value = "Abortos y legrados uterinos punibles."
    Range("F11").Value = "Enfermedades de transmisi�n sexual, excepto el S�ndrome de Inmunodeficiencia Adquirida (SIDA)."
    Range("F12").Value = "En ning�n caso se cubrir�n la renta o compra aparatos auditivos."
    Range("F13").Value = "Cualquier gasto realizado fuera de la vigencia de la p�liza."
    
    Range("F15").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub





