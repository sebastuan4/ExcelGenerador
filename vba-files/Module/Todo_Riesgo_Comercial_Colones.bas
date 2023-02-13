Attribute VB_Name = "Todo_Riesgo_Comercial_Colones"
Sub OCEANICA()
    Range("B1").Value = "TODO RIESGO INDUSTRIAL Y COMERCIAL COLONES"
    Range("B2").Value = "A: DA�OS DIRECTOS A LAS PROPIEDADES"
    Range("B3").Value = "B: ROTURA DE MAQUINARIAS Y EQUIPOS ELECTR�NICOS"
    Range("B4").Value = "C: LUCRO CESANTE"

    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    
    Range("B6").Value = "Condiciones Particulares"
    Range("B7").Value = "Inserte Condiciones Particulares"
    
    Range("B9").Value = "Condiciones Generales"
    Range("B10").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOcHOmK_EDiSWHd1Ig?e=uifFlc"
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Cualquier responsabilidad civil extracontractual o contractual."
    Range("F3").Value = "Actos de fraude, deshonestidad e infidelidad o actos intencionales del Asegurado o cualquiera de sus empleados, p�rdidas o faltantes descubiertos al practicar un inventario normal."
    Range("F4").Value = "Acci�n de ratas, comej�n, gorgojos, polillas o de cualquier animal en general y, en particular, los que puedan considerarse como plagas, germinaci�n de semillas o cultivos."
    Range("F5").Value = "Acciones u omisiones del Asegurado, sus empleados o personas actuando en su representaci�n o a quienes se les haya encargado la custodia de los bienes asegurados, que a criterio del instituto produzcan o agraven las p�rdidas."
    Range("F6").Value = "Desgaste o deterioro paulatino como consecuencia del uso o funcionamiento normal, erosi�n, corrosi�n, oxidaci�n, cavitaci�n, herrumbre."
    Range("F7").Value = "Saqueo, excepto si el siniestro ocurrido es a consecuencia de un evento amparado en la p�liza."
    Range("F8").Value = "Defectos o vicios propios ya existentes en la maquinaria o equipo electr�nico al iniciarse el seguro, de los cuales tenga conocimiento el Asegurado, sus representantes o personas responsables de la direcci�n t�cnica."
    Range("F9").Value = "Desgaste o deterioro paulatino como consecuencia del uso o funcionamiento normal, erosi�n, corrosi�n, oxidaci�n, cavitaci�n, herrumbre."
    Range("F10").Value = "Cualquier p�rdida consecuencial o p�rdidas indirectas o remotas que no se ajusten de un todo a las condiciones descritas en este amparo."
    Range("F11").Value = "Alguna ordenanza, local o estatal, o influencia de alguna ley reguladora de las construcciones o reparaciones de edificios o estructuras."
    
    Range("F13").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
