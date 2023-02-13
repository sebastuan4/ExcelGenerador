Attribute VB_Name = "Autos"
Sub ins()
    Range("B1").Value = "AUTOMÓVILES"
    Range("B2").Value = "A: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL POR LESIÓN Y/O MUERTE DE PERSONAS."
    Range("B3").Value = "C: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL POR DAÑOS A LA PROPIEDAD DE TERCEROS."
    Range("B4").Value = "D: COLISIÓN Y/O VUELCO."
    Range("B5").Value = "F: ROBO Y/O HURTO."
    Range("B6").Value = "G: MULTIASISTENCIA AUTOMÓVILES."
    Range("B7").Value = "H: RIESGOS ADICIONALES."
    Range("B8").Value = "I: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EXTENDIDA POR LESIÓN Y/O MUERTE DE PERSONAS Y DAÑOS A LA PROPIEDAD DE TERCEROS."
    Range("B9").Value = "L: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EXTENDIDA POR LESIÓN Y/O MUERTE DE PERSONAS Y DAÑOS A LA PROPIEDAD DE TERCEROS POR EL USO DE UN AUTO SUSTITUTO."
    Range("B10").Value = "B: SERVICIOS MÉDICOS FAMILIARES BÁSICA."
    Range("B11").Value = "E: GASTOS LEGALES."
    Range("B12").Value = "J: PÉRDIDA DE OBJETOS PERSONALES."
    Range("B13").Value = "K: INDEMNIZACIÓN PARA TRANSPORTE ALTERNATIVO"
    Range("B14").Value = "M: MULTIASISTENCIA EXTENDIDA."
    Range("B15").Value = "N: EXENCIÓN DE DEDUCIBLE."
    Range("B16").Value = "P: SERVICIOS MÉDICOS FAMILIARES PLUS Y MUERTE DE LOS OCUPANTES DEL VEHÍCULO ASEGURADO."
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
    
    Range("B27").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El Asegurado incumpla con lo establecido en el Artículo “Obligaciones del Asegurado” de este Contrato."
    Range("F3").Value = "Actos malintencionados cometidos por parte del Asegurado, de sus empleados, el conductor o personas que actúen en su nombre o a la que se le haya confiado la custodia del vehículo."
    Range("F4").Value = "Las obligaciones, compromisos, arreglos, convenios sean éstos judiciales o extrajudiciales que contraiga el Asegurado."
    Range("F5").Value = "Los casos donde el conductor del vehículo asegurado no cuente con la licencia habilitante según definición de este Contrato."
    Range("F6").Value = "El uso del vehículo declarado en la Solicitud del Seguro ha sido variado en forma permanente o reiterada sin el debido consentimiento del Instituto, siempre que esa modificación implique una agravación del riesgo asegurado por la cobertura específica"
    Range("F7").Value = "El automóvil asegurado sea utilizado para el transporte privado de personas."
    Range("F8").Value = "Sea utilizado en competencias o en pruebas de seguridad."
    Range("F9").Value = "Haya sido puesto a disposición o uso de persona distinta del Asegurado, por contrato de arrendamiento, venta condicional, convenio o promesa de compra, prenda, gravamen o condición que no haya sido declarada en esta póliza."
    Range("F10").Value = "Pérdida de beneficios anticipada."
    Range("F11").Value = "Si al ocurrir un accidente, el Conductor del vehículo asegurado se encuentra bajo la influencia o efectos del alcohol, drogas tóxicas o perturbadoras, estupefacientes, sustancias psicotrópicas, estimulantes u otras sustancias."
    Range("F12").Value = "Para cobertura de daño directo: Los daños en la cabina de pasajeros, sus componentes y vidrios del automóvil asegurado, sean causados por bultos u otros objetos que sea transportado en dicha cabina. El daño que produzca al automóvil asegurado la carga transportada, e.    El daño que el remolque, el remolque liviano o la carreta produzca, al automóvil asegurado que realiza la acción de remolcar o halar."
    
    Range("F21").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub Lafise()
    Range("B1").Value = "AUTOMÓVILES"
    Range("B2").Value = "A: Responsabilidad Civil Extracontractual por Lesión y/o Muerte de Personas"
    Range("B3").Value = "B: Responsabilidad Civil Extracontractual por Daños a la Propiedad de Terceras Personas."
    Range("B4").Value = "C: Colisión y Vuelco"
    Range("B5").Value = "D: Responsabilidad Civil Extracontractual extendida."
    Range("B6").Value = "E: Gastos médicos por Lesión de familiares ocupantes del vehículo asegurado por accidente y pago de gastos funerarios."
    Range("B7").Value = "F: Robo y Hurto."
    Range("B8").Value = "G: Riesgos Adicionales."
    Range("B9").Value = "H: Equipo especial."
    Range("B10").Value = "J: DEDUCIBLE 0"
    Range("B11").Value = "K: Asistencia en Carreteras."
    Range("B12").Value = "L: Responsabilidad Civil Extracontractual bajo los efectos del Alcohol."
    Range("B13").Value = "M: Auto sustituto."
    Range("B14").Value = "N: Servicios dentales por accidente automovilístico."

    
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
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El vehículo asegurado, al momento del siniestro, no cuente con los requisitos exigidos por la Ley de Tránsito de la República de Costa Rica para su circulación en las vías nacionales (Revisión Técnica Vehicular vigente, Marchamo debidamente pagado)."
    Range("F3").Value = "Cuando el vehículo asegurado sea utilizado en actividades diferentes al uso declarado en la solicitud de seguro, sin que medie previa autorización por escrito de SEGUROS LAFISE."
    Range("F4").Value = "Los daños materiales del vehículo asegurado ni los ocasionados por el mismo a terceros, cuando el Tomador y/o Asegurado al momento del siniestro no esté al día en el pago de las primas, sea que no haya cancelado las primas totales o las fracciones de primas convenidas en las fechas establecidas."
    Range("F5").Value = "Cuando el conductor del vehículo asegurado se encuentre bajo los efectos del licor, o bajo la influencia de estupefacientes o drogas tóxicas; a excepción de que se hubiere suscrito la cobertura correspondiente para este riesgo. Esta exclusión no opera si el Tomador y/o Asegurado o el conductor al momento del siniestro es absuelto en sede judicial por esta situación."
    Range("F6").Value = "El uso del vehículo declarado en la Solicitud del Seguro ha sido variado en forma permanente o reiterada sin el debido consentimiento del Instituto, siempre que esa modificación implique una agravación del riesgo asegurado por la cobertura específica"
    Range("F7").Value = "Los daños, las pérdidas o las responsabilidades que sufra u ocasione el vehículo asegurado, mientras el vehículo esté tomando parte, directa o indirectamente, en cualquier actividad ilícita, carreras, pruebas o contiendas de seguridad, resistencia o velocidad, al utilizarse para fines de enseñanza o de instrucción de su manejo o funcionamiento, empuje, remolque de otro vehículo, o para transporte de pasajeros mediante remuneración monetaria o de cualquier otra clase, tratándose de automóviles para usos particulares."
    Range("F8").Value = "El vehículo asegurado sea utilizado en la organización, ejecución o represión de huelga, paro, disturbio, motín, así como hechos que alteren el orden público."
    Range("F9").Value = "Haya sido puesto a disposición o uso de persona distinta del Asegurado, por contrato de arrendamiento, venta condicional, convenio o promesa de compra, prenda, gravamen o condición que no haya sido declarada en esta póliza."
    Range("F10").Value = "Los casos donde el conductor del vehículo asegurado sea conducido por persona que carezca de licencia que le autorice manejar la categoría del vehículo asegurado, o la tenga vencida, o aun contando con el Permiso Temporal de Aprendizaje emitido por el Ministerio de Obras Públicas y Transportes, no cumpla con la normativa que autoriza su utilización."
    Range("F11").Value = "Si al ocurrir un accidente, el Conductor del vehículo asegurado se encuentra bajo la influencia o efectos del alcohol, drogas tóxicas o perturbadoras, estupefacientes, sustancias psicotrópicas, estimulantes u otras sustancias."
    
    Range("F22").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
Sub QUALITAS()
    Range("B1").Value = "AUTOMÓVILES"
    Range("B2").Value = "1. DANOS MATERIALES"
    Range("B3").Value = "1.1 Rotura de Cristales"
    Range("B4").Value = "1.2 Riesgos Adicionales"
    Range("B5").Value = "2. Robo Total"
    Range("B6").Value = "3.1 RC Personas"
    Range("B7").Value = "3.2 RC Bienes"
    Range("B8").Value = "3.3 RC Complementaria"
    Range("B9").Value = "3.6 RC Daños a Ocupantes"
    Range("B10").Value = "4 Daños Legales"
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
    
    Range("B21").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
End Sub

Sub OCEANICA()
    Range("B1").Value = "AUTOMÓVILES"
    Range("B2").Value = "A: RESPONSABILIDAD CIVIL (COBERTURA BÁSICA)"
    Range("B3").Value = "D: DAÑO DIRECTO POR COLISIÓN Y/O VUELCO "
    Range("B4").Value = "F: ROBO Y/O HURTO "
    Range("B5").Value = "H: RIESGOS ADICIONALES."
    Range("B6").Value = "G: BENEFICIOS Y ASISTENCIAS"
    Range("B7").Value = "B: ATENCIÓN MÉDICA Y GASTOS FUNERARIOS"
    Range("B8").Value = "C: RESPONSABILIDAD CIVIL BAJO LOS EFECTOS DEL ALCOHOL"
    Range("B9").Value = "E: PARQUEO SEGURO"
    Range("B10").Value = "J: PÉRDIDA O SUSTRACCIÓN DE EFECTOS PERSONALES "
    Range("B11").Value = "P - PÉRDIDA TOTAL "
    Range("B12").Value = "K: SUSTITUCIÓN VEHÍCULO"
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
    
    Range("B22").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "El conductor no cuenta con licencia habilitante para el tipo de vehículo o que la misma no fue renovada dentro de los tres meses siguientes a su expiración, o bien esté suspendida."
    Range("F3").Value = "El vehículo no cuenta con su derecho de circulación o revisión técnica (RTV) al día."
    Range("F4").Value = "Cuando Tomador y/o asegurado asuma la responsabilidad ante sede judicial sin consentimiento expreso por parte de OCEÁNICA."
    Range("F5").Value = "Cuando al momento de la ocurrencia del evento, no se cuente con interés asegurable"
    Range("F6").Value = "Cuando el bien asegurado haya sido puesto a disposición o uso de una persona distinta del asegurado nombrado por medio de un contrato de arrendamiento, venta condicional, promesa de compra, o condición que no haya sido declarada en la póliza. "
    Range("F7").Value = "Comprobado exceso de velocidad temeraria a más de 150 km/h. "
    Range("F8").Value = "Transporte de bultos u objetos dentro de la cabina que produzcan daños internos             "
    Range("F9").Value = "Abandono del Vehículo Asegurado, su inmersión o tránsito voluntario en lugares inundados (lagos, carreteras, puentes, playas, esteros, bahías) o conducirlo en carreteras o lugares no viables para ello."
    Range("F10").Value = "Cuando el vehículo sea conducido por persona que se encuentre bajo los efectos de drogas o sustancias que produzcan estados de alteración para la adecuada conducción vehicular de conformidad con los parámetros establecidos en la Ley 9078 y Código Penal"
    
    Range("F22").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
