Attribute VB_Name = "Hogar"
Sub ins()
    Range("B1").Value = "Hogar Comprensivo"
    Range("B2").Value = "V:  Da�o Directo de Bienes Inmuebles"
    Range("B3").Value = "Y:  Da�o Directo de Contenidos (ampara robo)"
    Range("B4").Value = "X:  Da�o Directo de Contenidos (excluye robo)"
    Range("B5").Value = "D:  Convulsiones de la Naturaleza"
    Range("B6").Value = "H:  P�rdida de Rentas por Contrato de Arrendamiento"
    Range("B7").Value = "K:  Responsabilidad Civil"
    Range("B8").Value = "M:  Riesgos del Trabajo Hogar"
    Range("B9").Value = "P:  Accidentes Personales"
    
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
    Range("B16").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOov7rOQdn6fubZgww?e=rrwfhn"
    
    Range("B19").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Reacci�n nuclear, irradiaci�n nuclear o contaminaci�n radiactiva por combustibles nucleares o desechos radiactivos, debidos a su propia combusti�n."
    Range("F3").Value = "Contaminaci�n."
    Range("F4").Value = "Fermentaci�n, vicio propio o combusti�n espont�nea, o por procedimientos de calefacci�n o desecaci�n, al cual hubiese sido sometida."
    Range("F5").Value = "P�rdidas consecuenciales de cualquier �ndole, salvo que cuente con la cobertura respectiva."
    Range("F6").Value = "Colillas de cigarrillo o similares, a menos que produzcan incendio."
    Range("F7").Value = "Por polvo o arena, sean o no tra�dos por el viento."
    Range("F8").Value = "Deslizamiento de rellenos en laderas"
    Range("F9").Value = "Ca�da, volteo o derrame de recipientes, tanques o dep�sitos que no contengan agua."
    Range("F10").Value = "P�rdida de beneficios anticipada."
    Range("F11").Value = "Uso il�cito del inmueble asegurado, o contrario a la actividad declarada en el contrato p�liza."
    Range("F12").Value = "Da�os causados por filtraciones de agua en paredes, muros, cubiertas de techos y pisos, por falta de mantenimiento preventivo y correctivo."
    Range("F13").Value = "Dolo."
    
    Range("F18").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
    'Formato de todo
End Sub
Sub OCEANICA()
    Range("B1").Value = "Hogar Integral"
    Range("B2").Value = "A: INCENDIO Y OTROS DA�OS"
    Range("B3").Value = "B: DESLIZAMIENTO, INUNDACI�N Y VIENTOS HURACANADOS"
    Range("B4").Value = "C: TEMBLOR, TERREMOTO, MAREMOTO Y ERUPCI�N VOLC�NICA"
    Range("B5").Value = "D: DA�OS EN TUBER�AS Y SIMILARES"
    Range("B6").Value = "E: MOT�N, CONMOCI�N CIVIL, DISTURBIOS POPULARES Y DA�OS MALICIOSOS"
    Range("B7").Value = "F: ROTURA DE VIDRIOS"
    Range("B8").Value = "G: DESPLAZAMIENTO TEMPORAL DEL CONTENIDO"
    Range("B9").Value = "H: INHABITABILIDAD DE LA VIVIENDA"
    Range("B10").Value = "I: ROBO"
    Range("B11").Value = "J: RESPONSABILIDAD CIVIL FAMILIAR"
    Range("B12").Value = "K: P�RDIDA DE RENTAS"
    Range("B13").Value = "L: MULTIASISTENCIA RESIDENCIAL"
    
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
    Range("C11").Value = "No contratada"
    Range("C12").Value = "No contratada"
    Range("C13").Value = "No contratada"
    
    Range("B15").Value = "Condiciones Particulares"
    Range("B16").Value = "Inserte Condiciones Particulares"
    
    Range("B18").Value = "Condiciones Generales"
    Range("B19").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOwYbxB3og5OYks-Kg?e=17Pjd9"
    
    Range("B20").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Concusi�n, a menos que sea causada por una explosi�n."
    Range("F3").Value = "Colisi�n donde intervenga un veh�culo conducido o manipulado por cualquier ocupante de la casa, o por cualquier persona que trabaje o resida con el Asegurado."
    Range("F4").Value = "Los da�os provocados por un incendio causado por dolo o culpa grave del Asegurado."
    Range("F5").Value = "Los da�os causados por fuego no hostil."
    Range("F6").Value = "Los fen�menos resultantes de sobre voltaje o sobre corriente, recalentamiento, corto circuito, perforaci�n o carbonizaci�n del aislamiento, lo mismo que chisporroteos, arcos voltaicos y arcos el�ctricos, a no ser que produzcan incendio."
    Range("F7").Value = "Inundaciones que tengan origen en fallas o falta de capacidad de los sistemas  de evacuaci�n de aguas residuales o pluviales de la vivienda asegurada y/o sus predios."
    Range("F8").Value = "Deslizamiento de rellenos en laderas"
    Range("F9").Value = "Fallas en los muros de contenci�n por falta de capacidad soportante."
    Range("F10").Value = "El deslizamiento de rellenos en laderas."
    Range("F11").Value = "Roturas producto de la ocurrencia de un evento amparado en otra cobertura, est� o no incluida en la p�liza."
    
    Range("F20").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
