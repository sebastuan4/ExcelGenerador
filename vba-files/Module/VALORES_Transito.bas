Attribute VB_Name = "VALORES_Transito"
Sub ins()
    Range("B1").Value = "VALORES EN TRÁNSITO COLONES"
    Range("B2").Value = "A: Valores en Tránsito"
    Range("B3").Value = "C: Transporte y Pago de Planillas"
    Range("B4").Value = "E: Agentes Vendedores y/o Cobradores"
    Range("B5").Value = "F: Caja Fuerte y/o Bóveda"
    Range("B6").Value = "G: Cajeros y/ Cajas Registradoras"
    Range("B7").Value = "H: Cajero Automático"
    Range("B8").Value = "I: Buzón Nocturno"
    Range("B9").Value = "L: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EXTENDIDA POR LESIÓN Y/O MUERTE DE PERSONAS Y DAÑOS A LA PROPIEDAD DE TERCEROS POR EL USO DE UN AUTO SUSTITUTO."
    Range("B10").Value = "J: Caja Chica"
    
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
    Range("B16").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOV8cjtccs05fD5hLA?e=ElAKx9"
    
    Range("B18").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el año póliza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las más actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Reacción y/o fisión y/o fusión y/o irradiación nuclear, contaminación radioactiva por combustibles nucleares o desechos radiactivos, debidos a su propia combustión."
    Range("F3").Value = "Defraudación y/o estafa, y/o chantaje faltante de liquidación, y/o faltantes de caja, y/o faltantes de mercaderías.  "
    Range("F4").Value = "Terremoto, temblor, erupción volcánica, tifón, huracán, tornado o ciclón. Esta exclusión no aplica a los riesgos cubiertos por la Cobertura F.  "
    Range("F5").Value = "Hurto. "
    Range("F6").Value = "Infidelidad de los empleados del Asegurado y/o sus representantes. "
    Range("F7").Value = "Cuando se autoricen para el transporte o custodia de los bienes, personas con antecedentes penales por delitos contra la propiedad.  "
    Range("F8").Value = "Daños o pérdidas consecuenciales de cualquier tipo, incluyendo las pérdidas de beneficios o lucro cesante o de ganancias producidas como consecuencia del siniestro.  "
    Range("F9").Value = "Acciones u omisiones del Asegurado, sus empleados o personas actuando en su representación o a quienes se les haya encargado la custodia de los valores, que a criterio del instituto produzcan o agraven las pérdidas.  "
    Range("F10").Value = "Transporte, custodia o manipulación de valores por personas menores de edad, o por personas sin relación laboral con el Asegurado o por empleados que no cuenten con autorización para dicha función. "
    Range("F11").Value = "Desvíos o interrupciones del trayecto que incrementen el riesgo cubierto. No aplica esta exclusión cuando el desvío o interrupción sean para evitar o disminuir el riesgo o para culminar el transporte de los valores asegurados. "
    
    Range("F18").Value = "La información suministrada es un resumen, con lo que su asesor considera es lo más importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
