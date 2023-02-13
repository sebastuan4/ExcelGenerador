Attribute VB_Name = "Transporte_Mercancias"
Sub ins()
    Range("B1").Value = "MULTIRIESGO COBERTURAS"
    Range("B2").Value = "H: RIESGOS DEL MEDIO DE TRANSPORTE."
    Range("B3").Value = "E: HUELGA."
    Range("B4").Value = "I:�ROBO Y/O ASALTO."
    Range("B5").Value = "J: MANIOBRAS DE CARGA Y DESCARGA."
    Range("B6").Value = "K:�MOVIMIENTOS BRUSCOS."
    Range("B7").Value = "L: CA�DA, COLISI�N O VUELCO DE MERCANC�AS."
    Range("B8").Value = "N: CA�DA DE MERCANC�A EN PREDIOS."
    Range("B9").Value = "P: FALLAS MEC�NICAS EN EL SISTEMA DE REFRIGERACI�N."
    Range("B10").Value = "Q: RESPONSABILIDAD CIVIL DERIVADA DE LA CARGA TRANSPORTADA POR VIA TERRESTRE."
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    Range("C10").Value = "No contratada"
    
    Range("B12").Value = "Condiciones Particulares"
    Range("B13").Value = "Inserte Condiciones Particulares"
    
    Range("B15").Value = "Condiciones Generales"
    Range("B16").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihOMwgO5PcLIJeu54CA?e=Ccnha9"
    
    Range("B18").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Derrames ordinarios, p�rdidas de peso o de volumen por merma natural y uso o desgaste de los bienes asegurados."
    Range("F3").Value = "Combusti�n espont�nea."
    Range("F4").Value = "Oxidaci�n, p�rdida de potencial de germinaci�n, evaporaci�n o alteraci�n qu�mica y putrefacci�n, que da�en la mercanc�a transportada."
    Range("F5").Value = "Se excluyen los da�os causados cuando el (los) conductor (es) del veh�culo que transporta la mercanc�a, se encuentre (n) en estado de ebriedad."
    Range("F6").Value = "La utilizaci�n de medios de transporte o contenedores excediendo la capacidad de remolque o carga recomendados por el fabricante, o cuyas caracter�sticas t�cnicas, tales como velocidad, tipo de contenedor, material de fabricaci�n, sistema de refrigeraci�n o dispositivos de manipulaci�n, no permitan un transporte seguro."
    Range("F7").Value = "P�rdidas, da�os o gastos consecuenciales que se originen por p�rdida de mercado o demora (que est� bajo control del Asegurado); aun cuando sean originadas por un riesgo cubierto."
    Range("F8").Value = "P�rdidas o da�os de combustibles."
    Range("F9").Value = "Si el conductor permite o favorece el transporte o ingreso al veh�culo o contenedor de personas no relacionadas con la empresa de transportes o Asegurado; siempre que este acto contribuya a la ocurrencia del siniestro."
    Range("F10").Value = "Colisi�n del medio de transporte y/o contenedor, contra la parte superior de los t�neles, cuya altura m�xima sea inferior a la altura del medio de transporte o contenedor."
    Range("F11").Value = "Insolvencia o fallo financiero del Asegurado u otro transportista.  "
    Range("F12").Value = "Vicio propio del objeto asegurado."
    Range("F13").Value = "Acciones u omisiones del Asegurado o el propietario de la mercanc�a, sus empleados o personas actuando en su representaci�n o a quienes se les haya encargado la custodia de las mercanc�as, que a criterio del Instituto produzcan o agraven las p�rdidas."
    
    Range("F18").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
