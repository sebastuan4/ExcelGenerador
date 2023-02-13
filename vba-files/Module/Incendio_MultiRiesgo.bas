Attribute VB_Name = "Incendio_MultiRiesgo"
Sub ins()
    Range("B1").Value = "Incendio Multiriesgo"
    Range("B2").Value = "T: TODO RIESGO NO CATASTR�FICOS"
    Range("B3").Value = "C: TODO RIESGO INUNDACI�N, DESLIZAMIENTO Y VIENTOS"
    Range("B4").Value = "D: TODO RIESGO CONVULSIONES DE LA NATURALEZA"
    Range("B5").Value = "F: P�RDIDA DE BENEFICIOS"
    Range("B6").Value = "H: P�RDIDA DE RENTA POR CONTRATO DE ARRENDAMIENTO"
    Range("B7").Value = "M: MANIOBRAS DE CARGA Y DESCARGA DE MERCADER�AS"
    Range("B8").Value = "Q: GASTOS EXTRA"
    Range("B9").Value = "R: ROBO O TENTATIVA DE ROBO"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    
    Range("B12").Value = "Condiciones Particulares"
    Range("B13").Value = "Inserte Condiciones Particulares"
    
    Range("B15").Value = "Condiciones Generales"
    Range("B16").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNo9DOVBOPpUD6AhIA?e=TaLRBo"
    
    Range("B19").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Las p�rdidas causadas por el Asegurado o sus representantes con la finalidad de obtener su propio beneficio."
    Range("F3").Value = "Riesgo de Transporte, a menos que se contrate la cobertura adicional de Transporte Interior."
    Range("F4").Value = "Todo lo relacionado con guerras, armas nucleares, reacci�n nuclear, actos de las autoridades."
    Range("F5").Value = "Cierre total o parcial del servicio, falta de ocupaci�n o suspensi�n de las actividades por un per�odo mayor de un mes, de las edificaciones aseguradas o que contengan los bienes asegurados."
    Range("F6").Value = "Desaparici�n misteriosa, faltante o merma de los bienes asegurados."
    Range("F7").Value = "Hurto o saqueo, salvo el que se produzca durante el incendio."
    Range("F8").Value = "Dolo y/o fraude del Asegurado y/o Tomador."
    Range("F9").Value = "Retraso, p�rdida de mercado u otros da�os consecuenciales."
    Range("F10").Value = "P�rdida de beneficios anticipada."
    Range("F11").Value = "Los da�os que sufra un bien asegurado, despu�s de que por un siniestro haya sido reparado provisionalmente por el Asegurado, hasta tanto la reparaci�n se haga en forma definitiva."
    Range("F12").Value = "Da�os sufridos a la propiedad asegurada cuando est� siendo objeto de trabajos a prueba, reparaci�n, ajuste, servicio u operaci�n de mantenimiento."
    Range("F13").Value = "P�rdidas y da�os provocados por inquilinos del Asegurado."
    
    Range("F18").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub OCEANICA()
    Range("B1").Value = "MULTIRIESGO COBERTURAS"
    Range("B2").Value = "A DA�OS DIRECTOS A LAS PROPIEDADES"
    Range("B3").Value = "COBERTURA B: ROTURA DE MAQUINARIAS Y EQUIPOS ELECTR�NICOS"
    Range("B4").Value = "COBERTURA C: LUCRO CESANTE"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    
    Range("B8").Value = "Condiciones Particulares"
    Range("B9").Value = "Inserte Condiciones Particulares"
    
    Range("B11").Value = "Condiciones Generales"
    Range("B12").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNpcuTt7_wa0V349-w?e=aCLWj8"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Nacionalizaci�n, confiscaci�n, incautaci�n, requisa, confiscaci�n, embargo, secuestro, expropiaci�n, destrucci�n o da�o por orden de cualquier gobierno o autoridad p�blica legalmente."
    Range("F3").Value = "Mermas y da�os por encogimiento, evaporaci�n, p�rdida de peso, contaminaci�n, cambios de color, sabor, textura, acabado, propiedades, acci�n de la luz y ralladuras, p�rdida de mercado, deterioro gradual, defecto latente, humedad o sequedad atmosf�rica, o cambios de temperatura."
    Range("F4").Value = "Acci�n de ratas, comej�n, gorgojos, polillas o de cualquier animal en general y, en particular, los que puedan considerarse como plagas, germinaci�n de semillas o cultivos."
    Range("F5").Value = "Actos de fraude, deshonestidad e infidelidad o actos intencionales del Asegurado o cualquiera de sus empleados, p�rdidas o faltantes descubiertos al practicar un inventario normal."
    Range("F6").Value = "Desaparici�n misteriosa, faltante o merma de los bienes asegurados."
    Range("F7").Value = "Cualquier responsabilidad civil extracontractual o contractual."
    Range("F8").Value = "Hundimientos, desplazamientos, agrietamientos o asentamientos, contracci�n o expansi�n de fundiciones, cimientos, muros, pisos, techos y pavimentos, excepto por terremoto."
    Range("F9").Value = "Hurto y/o desaparici�n misteriosa."
    Range("F10").Value = "Vicio propio."
    Range("F11").Value = "P�rdidas indirectas y/o consecuenciales."
    Range("F12").Value = "El da�o ocasionado a cualquier aparato el�ctrico o parte de la instalaci�n el�ctrica, causado por corriente el�ctrica generada artificialmente, a menos que se produzca incendio, en cuyo caso OCE�NICA solo est� obligada a pagar las p�rdidas o da�os causados por dicho incendio."

    Range("F16").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
