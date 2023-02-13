Attribute VB_Name = "RC"
Sub ins()
    Range("B1").Value = "RESPONSABILIDAD CIVIL"
    Range("B2").Value = "Coberturas"
    Range("B3").Value = "L: RESPONSABILIDAD CIVIL."
    Range("B4").Value = "M: RESPONSABILIDAD CIVIL PRODUCTOS."
    Range("B5").Value = "N: RESPONSABILIDAD CIVIL PATRONAL."
    Range("B6").Value = "O: RESPONSABILIDAD CIVIL COLISIÓN Y/O"
    Range("B7").Value = "P: RESPONSABILIDAD CIVIL ROBO DE VEHÍCULOS."
    Range("B8").Value = "Q: RESPONSABILIDAD CIVIL PRUEBA VEHÍCULOS PARA TALLERES."
    Range("B9").Value = "R: RESPONSABILIDAD CIVIL ATENCION MÉDICA INMEDIATA/"
    
    Range("C2").Value = "Deducibles"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    Range("C9").Value = "No contratada"
    
    Range("B11").Value = "Condiciones Particulares"
    Range("B12").Value = "Inserte Condiciones Particulares"
    
    Range("B14").Value = "Condiciones Generales"
    Range("B15").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihNxVB-GuvFeev2NyWQ?e=W06kaD"
    
    Range("B17").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Guerra, invasión, actos de enemigos extranjeros, actividades u operaciones militares."
    Range("F3").Value = "Reacción nuclear, irradiación nuclear o contaminación radiactiva por combustibles nucleares o desechos radiactivos."
    Range("F4").Value = "Actos deliberadamente perjudiciales, actos mal intencionados o cometidos con dolo por parte del Asegurado y/o Tomador."
    Range("F5").Value = "Reclamaciones de la que el Asegurado y/o Tomador hubiera tenido conocimiento en el momento de formalizar el contrato."
    Range("F6").Value = "Eventos de la naturaleza."
    Range("F7").Value = "Responsabilidad Civil Contractual."
    Range("F8").Value = "Los daños derivados del indebido ejercicio profesional del Asegurado."
    Range("F9").Value = "Reclamaciones y Demandas provenientes del Exterior."
    Range("F10").Value = "Las obligaciones legalmente imputables al Asegurado bajo la  Legislación de Riesgos del Trabajo"
    

    Range("F12").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

Sub OCEANICA()
    Range("B1").Value = "RESPONSABILIDAD CIVIL"
    Range("B2").Value = "Coberturas"
    Range("B3").Value = "A: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL"
    Range("B4").Value = "F: RESP CIVIL POR EL USO DE PARQUEOS"
    Range("B5").Value = "B: ATENCIÓN MÉDICA"
    Range("B6").Value = "C:LAVANDERÍAS Y GUARDARROPAS"
    Range("B7").Value = "D: EQUIPAJE DE HUÉSPEDES"
    Range("B8").Value = "E: BIENES RESGUARDADOS EN CAJAS DE SEGURIDAD"
    
    Range("C2").Value = "Deducibles"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    Range("C8").Value = "No contratada"
    
    Range("B10").Value = "Condiciones Particulares"
    Range("B11").Value = "Inserte Condiciones Particulares"
    
    Range("B13").Value = "Condiciones Generales"
    Range("B14").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOsOPPDbDxcmHh-Mmw?e=hkco3V"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Riesgos atómicos y nucleares, salvo empleo autorizado en la medicina."
    Range("F3").Value = "Riesgos relacionados a la navegación aérea y espacial, a productos para aeronavesicos y nucleares, salvo empleo autorizado en la medicina."
    Range("F4").Value = "La Responsabilidad Civil que tenga su origen en pérdidas financieras puras (daños patrimoniales sin daño físico)."
    Range("F5").Value = "Daños causados por contaminación paulatina."
    Range("F6").Value = "Pólizas de cumplimiento o garantía."
    Range("F7").Value = "Coberturas retroactivas de riesgos del pasado."
    Range("F8").Value = "Coberturas de Retirada de productos."
    Range("F9").Value = "Responsabilidad Civil Profesional (incluyendo errores y omisiones)."
    Range("F10").Value = "Productos farmacéuticos."
    Range("F11").Value = "Acto deliberado."
    Range("F12").Value = "Daños o lesiones causados por mercancías vendidas ."
    

    Range("F16").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
Sub Lafise()
    Range("B1").Value = "RESPONSABILIDAD CIVIL"
    Range("B2").Value = "Coberturas"
    Range("B3").Value = "A: Responsabilidad Civil Extracontractual Subjetiva"
    Range("B4").Value = "B: Atención Médica Inmediata"
   
    Range("C2").Value = "Deducibles"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    
    Range("B10").Value = "Condiciones Particulares"
    Range("B11").Value = "Inserte Condiciones Particulares"
    
    Range("B13").Value = "Condiciones Generales"
    Range("B14").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihPpL5pz9OzNBGwCTNg?e=JhhxRQ"
    
    Range("B16").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Toda responsabilidad imputable al Asegurado de acuerdo con la legislación de Riesgos del Trabajo o cualquier otra disposición legal conexa."
    Range("F3").Value = "Daños consecuenciales derivado de cualquier riesgo cubierto por esta póliza. "
    Range("F4").Value = "Responsabilidad Civil Cruzada. "
    Range("F5").Value = "Responsabilidad Civil Objetiva. "
    Range("F6").Value = "Responsabilidad Profesional del Asegurado. "
    Range("F7").Value = "Responsabilidad Penal del Asegurado o sus representantes."
    Range("F8").Value = "Daños por productos u objetos cuya fabricación, entrega, ejecución, carecen de permiso o licencias respectivas. "
    Range("F9").Value = "Responsabilidad Civil Profesional (incluyendo errores y omisiones)."
    Range("F10").Value = "Cualquier Responsabilidad Civil imputable al Asegurado, cuando esté realizando actividades no declaradas en las Condiciones Particulares. "
    Range("F11").Value = "Pérdidas o daños de los bienes personales que se encuentren dentro de vehículos"
    Range("F12").Value = "Daños que sufran o causen automóviles propiedad del Tomador y/o Asegurado o de sus empleados. "
    

    Range("F16").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
Sub Mapfre()
    Range("B1").Value = "Responsabilidad Civil Coberturas"
    Range("B2").Value = "A: Responsabilidad Civil Extracontractual y Subjetiva (Básica)"
    Range("B3").Value = "B: Gastos Médicos de Urgencia"
    Range("B4").Value = "C: RC Extracontractual en Lavandería y Guardarropa de Hoteles y Similares."
    Range("B5").Value = "D: RC en Cajitas de seguridad de Hoteles y similares"
    Range("B6").Value = "E: RC por Equipajes de huéspedes en Hoteles y Similares"
    Range("B7").Value = "F: RC Extracontractual por el uso del Parqueo Brindado por el Asegurado"
    
    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    Range("C6").Value = "No contratada"
    Range("C7").Value = "No contratada"
    
    Range("B9").Value = "Condiciones Particulares"
    Range("B10").Value = "Inserte Condiciones Particulares"
    
    Range("B12").Value = "Condiciones Generales"
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihPtyIuCr0H_ErfOqYg?e=g1Csjm"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Las lesiones, muertes o daños ocasionados a quien no sea tercero conforme se define en esta póliza."
    Range("F3").Value = "Daños derivados de cualquier responsabilidad civil profesional del Asegurado y/o sus empleados."
    Range("F4").Value = "Responsabilidad derivada de actividades de riesgo portuario o aeroportuario."
    Range("F5").Value = "Lesiones y/o muerte a personas y/o daños y perjuicios provocados por la culpa inexcusable o por responsabilidad atribuible al tercero afectado."
    Range("F6").Value = "Daños al medio ambiente."
    Range("F7").Value = "Explotación y producción de petróleo o derivados."
    Range("F8").Value = "Aguas negras, basura o sustancias residuales, sean industriales o residenciales."
    Range("F9").Value = "Reclamaciones de las que el asegurado hubiere tenido conocimiento en el momento de formalizar el contrato."
    Range("F10").Value = "Reclamaciones y Demandas tramitadas, juzgadas o provenientes de cualquier País Extranjero."
    Range("F11").Value = "Contaminación gradual, paulatina, lenta, progresiva o crónica."
    Range("F12").Value = "Eventos de la naturaleza."
    Range("F13").Value = "Operaciones que hayan sido definitivamente terminadas o abandonadas por el asegurado."
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

