Attribute VB_Name = "Autos"
Sub ins()
    Range("B1").Value = "AUTOM�VILES"
    Range("B2").Value = "A: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL POR LESI�N Y/O MUERTE DE PERSONAS."
    Range("B3").Value = "C: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL POR DA�OS A LA PROPIEDAD DE TERCEROS."
    Range("B4").Value = "D: COLISI�N Y/O VUELCO."
    Range("B5").Value = "F: ROBO Y/O HURTO."
    Range("B6").Value = "G: MULTIASISTENCIA AUTOM�VILES."
    Range("B7").Value = "H: RIESGOS ADICIONALES."
    Range("B8").Value = "I: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EXTENDIDA POR LESI�N Y/O MUERTE DE PERSONAS Y DA�OS A LA PROPIEDAD DE TERCEROS."
    Range("B9").Value = "L: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EXTENDIDA POR LESI�N Y/O MUERTE DE PERSONAS Y DA�OS A LA PROPIEDAD DE TERCEROS POR EL USO DE UN AUTO SUSTITUTO."
    Range("B10").Value = "B: SERVICIOS M�DICOS FAMILIARES B�SICA."
    Range("B11").Value = "E: GASTOS LEGALES."
    Range("B12").Value = "J: P�RDIDA DE OBJETOS PERSONALES."
    Range("B13").Value = "K: INDEMNIZACI�N PARA TRANSPORTE ALTERNATIVO"
    Range("B14").Value = "M: MULTIASISTENCIA EXTENDIDA."
    Range("B15").Value = "N: EXENCI�N DE DEDUCIBLE."
    Range("B16").Value = "P: SERVICIOS M�DICOS FAMILIARES PLUS Y MUERTE DE LOS OCUPANTES DEL VEH�CULO ASEGURADO."
    Range("B17").Value = "Y: EXTRATERRITORIALIDAD."
    Range("B18").Value = "Z: RIESGOS PARTICULARES."

    
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
    Range("C14").Value = "No contratada"
    Range("C15").Value = "No contratada"
    Range("C16").Value = "No contratada"
    Range("C17").Value = "No contratada"
    Range("C18").Value = "No contratada"
    
    Range("B21").Value = "Condiciones Particulares"
    Range("B22").Value = "Inserte Condiciones Particulares"
    
    Range("B24").Value = "Condiciones Generales"
    Range("B25").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNo9DOVBOPpUD6AhIA?e=TaLRBo"
    
    Range("B27").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El Asegurado incumpla con lo establecido en el Art�culo �Obligaciones del Asegurado� de este Contrato."
    Range("F3").Value = "Actos malintencionados cometidos por parte del Asegurado, de sus empleados, el conductor o personas que act�en en su nombre o a la que se le haya confiado la custodia del veh�culo."
    Range("F4").Value = "Las obligaciones, compromisos, arreglos, convenios sean �stos judiciales o extrajudiciales que contraiga el Asegurado."
    Range("F5").Value = "Los casos donde el conductor del veh�culo asegurado no cuente con la licencia habilitante seg�n definici�n de este Contrato."
    Range("F6").Value = "El uso del veh�culo declarado en la Solicitud del Seguro ha sido variado en forma permanente o reiterada sin el debido consentimiento del Instituto, siempre que esa modificaci�n implique una agravaci�n del riesgo asegurado por la cobertura espec�fica"
    Range("F7").Value = "El autom�vil asegurado sea utilizado para el transporte privado de personas."
    Range("F8").Value = "Sea utilizado en competencias o en pruebas de seguridad."
    Range("F9").Value = "Haya sido puesto a disposici�n o uso de persona distinta del Asegurado, por contrato de arrendamiento, venta condicional, convenio o promesa de compra, prenda, gravamen o condici�n que no haya sido declarada en esta p�liza."
    Range("F10").Value = "P�rdida de beneficios anticipada."
    Range("F11").Value = "Si al ocurrir un accidente, el Conductor del veh�culo asegurado se encuentra bajo la influencia o efectos del alcohol, drogas t�xicas o perturbadoras, estupefacientes, sustancias psicotr�picas, estimulantes u otras sustancias."
    Range("F12").Value = "Para cobertura de da�o directo: Los da�os en la cabina de pasajeros, sus componentes y vidrios del autom�vil asegurado, sean causados por bultos u otros objetos que sea transportado en dicha cabina. El da�o que produzca al autom�vil asegurado la carga transportada, e.    El da�o que el remolque, el remolque liviano o la carreta produzca, al autom�vil asegurado que realiza la acci�n de remolcar o halar."
    
    Range("F21").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub Lafise()
    Range("B1").Value = "AUTOM�VILES"
    Range("B2").Value = "A: Responsabilidad Civil Extracontractual por Lesi�n y/o Muerte de Personas"
    Range("B3").Value = "B: Responsabilidad Civil Extracontractual por Da�os a la Propiedad de Terceras Personas."
    Range("B4").Value = "C: Colisi�n y Vuelco"
    Range("B5").Value = "D: Responsabilidad Civil Extracontractual extendida."
    Range("B6").Value = "E: Gastos m�dicos por Lesi�n de familiares ocupantes del veh�culo asegurado por accidente y pago de gastos funerarios."
    Range("B7").Value = "F: Robo y Hurto."
    Range("B8").Value = "G: Riesgos Adicionales."
    Range("B9").Value = "H: Equipo especial."
    Range("B10").Value = "J: DEDUCIBLE 0"
    Range("B11").Value = "K: Asistencia en Carreteras."
    Range("B12").Value = "L: Responsabilidad Civil Extracontractual bajo los efectos del Alcohol."
    Range("B13").Value = "M: Auto sustituto."
    Range("B14").Value = "N: Servicios dentales por accidente automovil�stico."

    
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
    Range("C14").Value = "No contratada"

    
    Range("B16").Value = "Condiciones Particulares"
    Range("B17").Value = "Inserte Condiciones Particulares"
    
    Range("B19").Value = "Condiciones Generales"
    Range("B20").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihOZnFPT0owSDesJSrg?e=XYsqRU"
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El veh�culo asegurado, al momento del siniestro, no cuente con los requisitos exigidos por la Ley de Tr�nsito de la Rep�blica de Costa Rica para su circulaci�n en las v�as nacionales (Revisi�n T�cnica Vehicular vigente, Marchamo debidamente pagado)."
    Range("F3").Value = "Cuando el veh�culo asegurado sea utilizado en actividades diferentes al uso declarado en la solicitud de seguro, sin que medie previa autorizaci�n por escrito de SEGUROS LAFISE."
    Range("F4").Value = "Los da�os materiales del veh�culo asegurado ni los ocasionados por el mismo a terceros, cuando el Tomador y/o Asegurado al momento del siniestro no est� al d�a en el pago de las primas, sea que no haya cancelado las primas totales o las fracciones de primas convenidas en las fechas establecidas."
    Range("F5").Value = "Cuando el conductor del veh�culo asegurado se encuentre bajo los efectos del licor, o bajo la influencia de estupefacientes o drogas t�xicas; a excepci�n de que se hubiere suscrito la cobertura correspondiente para este riesgo. Esta exclusi�n no opera si el Tomador y/o Asegurado o el conductor al momento del siniestro es absuelto en sede judicial por esta situaci�n."
    Range("F6").Value = "El uso del veh�culo declarado en la Solicitud del Seguro ha sido variado en forma permanente o reiterada sin el debido consentimiento del Instituto, siempre que esa modificaci�n implique una agravaci�n del riesgo asegurado por la cobertura espec�fica"
    Range("F7").Value = "Los da�os, las p�rdidas o las responsabilidades que sufra u ocasione el veh�culo asegurado, mientras el veh�culo est� tomando parte, directa o indirectamente, en cualquier actividad il�cita, carreras, pruebas o contiendas de seguridad, resistencia o velocidad, al utilizarse para fines de ense�anza o de instrucci�n de su manejo o funcionamiento, empuje, remolque de otro veh�culo, o para transporte de pasajeros mediante remuneraci�n monetaria o de cualquier otra clase, trat�ndose de autom�viles para usos particulares."
    Range("F8").Value = "El veh�culo asegurado sea utilizado en la organizaci�n, ejecuci�n o represi�n de huelga, paro, disturbio, mot�n, as� como hechos que alteren el orden p�blico."
    Range("F9").Value = "Haya sido puesto a disposici�n o uso de persona distinta del Asegurado, por contrato de arrendamiento, venta condicional, convenio o promesa de compra, prenda, gravamen o condici�n que no haya sido declarada en esta p�liza."
    Range("F10").Value = "Los casos donde el conductor del veh�culo asegurado sea conducido por persona que carezca de licencia que le autorice manejar la categor�a del veh�culo asegurado, o la tenga vencida, o aun contando con el Permiso Temporal de Aprendizaje emitido por el Ministerio de Obras P�blicas y Transportes, no cumpla con la normativa que autoriza su utilizaci�n."
    Range("F11").Value = "Si al ocurrir un accidente, el Conductor del veh�culo asegurado se encuentra bajo la influencia o efectos del alcohol, drogas t�xicas o perturbadoras, estupefacientes, sustancias psicotr�picas, estimulantes u otras sustancias."
    
    Range("F22").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
Sub QUALITAS()
    Range("B1").Value = "AUTOM�VILES"
    Range("B2").Value = "1. DANOS MATERIALES"
    Range("B3").Value = "1.1 Rotura de Cristales"
    Range("B4").Value = "1.2 Riesgos Adicionales"
    Range("B5").Value = "2. Robo Total"
    Range("B6").Value = "3.1 RC Personas"
    Range("B7").Value = "3.2 RC Bienes"
    Range("B8").Value = "3.3 RC Complementaria"
    Range("B9").Value = "3.6 RC Da�os a Ocupantes"
    Range("B10").Value = "4 Da�os Legales"
    Range("B11").Value = "5 Gastos Medicos Ocupantes"
    Range("B12").Value = "15 Asistencia Vial"
    Range("B13").Value = "23 Robo Parcial"
    
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
    Range("B19").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihN49UsXfkNQnE6AwjQ?e=dTmkgk"
    
    Range("B21").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
End Sub

Sub OCEANICA()
    Range("B1").Value = "AUTOM�VILES"
    Range("B2").Value = "A: RESPONSABILIDAD CIVIL (COBERTURA B�SICA)"
    Range("B3").Value = "D: DA�O DIRECTO POR COLISI�N Y/O VUELCO "
    Range("B4").Value = "F: ROBO Y/O HURTO "
    Range("B5").Value = "H: RIESGOS ADICIONALES."
    Range("B6").Value = "G: BENEFICIOS Y ASISTENCIAS"
    Range("B7").Value = "B: ATENCI�N M�DICA Y GASTOS FUNERARIOS"
    Range("B8").Value = "C: RESPONSABILIDAD CIVIL BAJO LOS EFECTOS DEL ALCOHOL"
    Range("B9").Value = "E: PARQUEO SEGURO"
    Range("B10").Value = "J: P�RDIDA O SUSTRACCI�N DE EFECTOS PERSONALES "
    Range("B11").Value = "P - P�RDIDA TOTAL "
    Range("B12").Value = "K: SUSTITUCI�N VEH�CULO"
    Range("B13").Value = "M: EQUIPO ESPECIAL"
    Range("B14").Value = "N: EXTRATERRITORIALIDAD"
 
    
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
    Range("C14").Value = "No contratada"
    
    Range("B16").Value = "Condiciones Particulares"
    Range("B17").Value = "Inserte Condiciones Particulares"
    
    Range("B19").Value = "Condiciones Generales"
    Range("B20").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNt3ddeZX4TtFMvKqg?e=8cKLVN"
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El conductor no cuenta con licencia habilitante para el tipo de veh�culo o que la misma no fue renovada dentro de los tres meses siguientes a su expiraci�n, o bien est� suspendida."
    Range("F3").Value = "El veh�culo no cuenta con su derecho de circulaci�n o revisi�n t�cnica (RTV) al d�a."
    Range("F4").Value = "Cuando Tomador y/o asegurado asuma la responsabilidad ante sede judicial sin consentimiento expreso por parte de OCE�NICA."
    Range("F5").Value = "Cuando al momento de la ocurrencia del evento, no se cuente con inter�s asegurable"
    Range("F6").Value = "Cuando el bien asegurado haya sido puesto a disposici�n o uso de una persona distinta del asegurado nombrado por medio de un contrato de arrendamiento, venta condicional, promesa de compra, o condici�n que no haya sido declarada en la p�liza. "
    Range("F7").Value = "Comprobado exceso de velocidad temeraria a m�s de 150 km/h. "
    Range("F8").Value = "Transporte de bultos u objetos dentro de la cabina que produzcan da�os internos             "
    Range("F9").Value = "Abandono del Veh�culo Asegurado, su inmersi�n o tr�nsito voluntario en lugares inundados (lagos, carreteras, puentes, playas, esteros, bah�as) o conducirlo en carreteras o lugares no viables para ello."
    Range("F10").Value = "Cuando el veh�culo sea conducido por persona que se encuentre bajo los efectos de drogas o sustancias que produzcan estados de alteraci�n para la adecuada conducci�n vehicular de conformidad con los par�metros establecidos en la Ley 9078 y C�digo Penal"
    
    Range("F22").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
