Attribute VB_Name = "Incendio_MultiRiesgo"
Sub ins()
    Range("B1").Value = "Incendio Multiriesgo"
    Range("B2").Value = "T: TODO RIESGO NO CATASTRÓFICOS"
    Range("B3").Value = "C: TODO RIESGO INUNDACIÓN, DESLIZAMIENTO Y VIENTOS"
    Range("B4").Value = "D: TODO RIESGO CONVULSIONES DE LA NATURALEZA"
    Range("B5").Value = "F: PÉRDIDA DE BENEFICIOS"
    Range("B6").Value = "H: PÉRDIDA DE RENTA POR CONTRATO DE ARRENDAMIENTO"
    Range("B7").Value = "M: MANIOBRAS DE CARGA Y DESCARGA DE MERCADERÍAS"
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
    
    Range("B19").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Las pérdidas causadas por el Asegurado o sus representantes con la finalidad de obtener su propio beneficio."
    Range("F3").Value = "Riesgo de Transporte, a menos que se contrate la cobertura adicional de Transporte Interior."
    Range("F4").Value = "Todo lo relacionado con guerras, armas nucleares, reacción nuclear, actos de las autoridades."
    Range("F5").Value = "Cierre total o parcial del servicio, falta de ocupación o suspensión de las actividades por un período mayor de un mes, de las edificaciones aseguradas o que contengan los bienes asegurados."
    Range("F6").Value = "Desaparición misteriosa, faltante o merma de los bienes asegurados."
    Range("F7").Value = "Hurto o saqueo, salvo el que se produzca durante el incendio."
    Range("F8").Value = "Dolo y/o fraude del Asegurado y/o Tomador."
    Range("F9").Value = "Retraso, pérdida de mercado u otros daños consecuenciales."
    Range("F10").Value = "Pérdida de beneficios anticipada."
    Range("F11").Value = "Los daños que sufra un bien asegurado, después de que por un siniestro haya sido reparado provisionalmente por el Asegurado, hasta tanto la reparación se haga en forma definitiva."
    Range("F12").Value = "Daños sufridos a la propiedad asegurada cuando está siendo objeto de trabajos a prueba, reparación, ajuste, servicio u operación de mantenimiento."
    Range("F13").Value = "Pérdidas y daños provocados por inquilinos del Asegurado."
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub OCEANICA()
    Range("B1").Value = "MULTIRIESGO COBERTURAS"
    Range("B2").Value = "A DAÑOS DIRECTOS A LAS PROPIEDADES"
    Range("B3").Value = "COBERTURA B: ROTURA DE MAQUINARIAS Y EQUIPOS ELECTRÓNICOS"
    Range("B4").Value = "COBERTURA C: LUCRO CESANTE"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    
    Range("B8").Value = "Condiciones Particulares"
    Range("B9").Value = "Inserte Condiciones Particulares"
    
    Range("B11").Value = "Condiciones Generales"
    Range("B12").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNpcuTt7_wa0V349-w?e=aCLWj8"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Nacionalización, confiscación, incautación, requisa, confiscación, embargo, secuestro, expropiación, destrucción o daño por orden de cualquier gobierno o autoridad pública legalmente."
    Range("F3").Value = "Mermas y daños por encogimiento, evaporación, pérdida de peso, contaminación, cambios de color, sabor, textura, acabado, propiedades, acción de la luz y ralladuras, pérdida de mercado, deterioro gradual, defecto latente, humedad o sequedad atmosférica, o cambios de temperatura."
    Range("F4").Value = "Acción de ratas, comején, gorgojos, polillas o de cualquier animal en general y, en particular, los que puedan considerarse como plagas, germinación de semillas o cultivos."
    Range("F5").Value = "Actos de fraude, deshonestidad e infidelidad o actos intencionales del Asegurado o cualquiera de sus empleados, pérdidas o faltantes descubiertos al practicar un inventario normal."
    Range("F6").Value = "Desaparición misteriosa, faltante o merma de los bienes asegurados."
    Range("F7").Value = "Cualquier responsabilidad civil extracontractual o contractual."
    Range("F8").Value = "Hundimientos, desplazamientos, agrietamientos o asentamientos, contracción o expansión de fundiciones, cimientos, muros, pisos, techos y pavimentos, excepto por terremoto."
    Range("F9").Value = "Hurto y/o desaparición misteriosa."
    Range("F10").Value = "Vicio propio."
    Range("F11").Value = "Pérdidas indirectas y/o consecuenciales."
    Range("F12").Value = "El daño ocasionado a cualquier aparato eléctrico o parte de la instalación eléctrica, causado por corriente eléctrica generada artificialmente, a menos que se produzca incendio, en cuyo caso OCEÁNICA solo está obligada a pagar las pérdidas o daños causados por dicho incendio."

    Range("F16").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
