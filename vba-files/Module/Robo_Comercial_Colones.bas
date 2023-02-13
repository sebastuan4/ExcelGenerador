Attribute VB_Name = "Robo_Comercial_Colones"
Sub ins()
    Range("B1").Value = "ROBO COMERCIAL COBERTURAS"
    Range("B2").Value = "A: ROBO Y TENTATIVA DE ROBO"
    Range("B3").Value = "B: BIENES DEPOSITADOS EN EXTERIORES "
    Range("B4").Value = "D: BIENES DE TERCEROS "
    Range("B5").Value = "E: BIENES EN TRÁNSITO"
    Range("B6").Value = "F: TRASLADO DE BIENES"
    Range("B7").Value = "G: MULTIASISTENCIA COMERCIAL (PLAN TOTAL PLUS)"
    
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
    Range("B13").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOZz1RFun3-EHHcgOA?e=Umb3WR"
    
    Range("B15").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Conmociones civiles, motines, huelgas, guerras civiles, rebeliones, insurrecciones, revoluciones, ley marcial, poder militar usurpado, confiscación, requisa, nacionalización o destrucción ordenadas por el gobierno o por la autoridad."
    Range("F3").Value = "Incendio o explosión."
    Range("F4").Value = "Robo o tentativa de robo consecuencia de las propiedades radiactivas, tóxicas, explosivas o de otra naturaleza peligrosa, de unidades nucleares explosivas o de un componente nuclear de ella."
    Range("F5").Value = "Reembolsos por servicios que el Asegurado contrate por sus propios medios."
    Range("F6").Value = "Daños al Inmueble causados por cualquier tipo de plaga que le haya invadido, aún cuando atente contra la seguridad del propio local comercial asegurado."
    Range("F7").Value = "Saqueo, excepto si el siniestro ocurrido es a consecuencia de un evento amparado en la póliza."
    Range("F8").Value = "Daños causados por filtraciones de humedad en muros y tejados."
    Range("F9").Value = "Inundaciones, provenientes de riesgos no cubiertos en esta cobertura."
    Range("F10").Value = "En el servicio de conexión con la red de proveedores médicos No se cubrirá económicamente el costo de la consulta ni los honorarios por servicios."
    
    Range("F15").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub

