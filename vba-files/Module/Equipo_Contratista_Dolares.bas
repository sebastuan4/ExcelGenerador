Attribute VB_Name = "Equipo_Contratista_Dolares"
Sub Lafise()
    Range("B1").Value = "Equipo Contratista"
    Range("B2").Value = "A: Todo Riesgo de Equipo de Contratista"
    Range("B3").Value = "D: Responsabilidad Civil Extracontractual y Subjetiva, bajo la Modalidad de L�mite �nico Combinado y L�mite Agregado AnuaL"
    Range("B4").Value = "F: Riesgos Diversos"

    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    
    Range("B6").Value = "Condiciones Particulares"
    Range("B7").Value = "Inserte Condiciones Particulares"
    
    Range("B9").Value = "Condiciones Generales"
    Range("B10").Value = "https://1drv.ms/b/s!Au8GQldWcy2ihOII4s-zYNg63sTPBA?e=4fl6Tu"
    
    Range("B12").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "P�rdida o da�os por aver�a mec�nica internas o fallas el�ctrica, congelaci�n de refrigerante o de otros fluidos, lubricaci�n deficiente o falta de aceite o derefrigerante.'"
    Range("F3").Value = "El desgaste de bandas de transmisi�n de toda clase, brocas, taladros, cuchillas o dem�s herramientas de cortar"
    Range("F4").Value = "Da�os o p�rdidas que sean a consecuencia directa de influencia continua de la operaci�n"
    Range("F5").Value = "Da�os ocasionados a consecuencia de la ejecuci�n de maniobras de carga y descarga. A menos que se haya contratado la Cobertura de Responsabilidad Civil de Operaciones de Carga y Descarga. "
    Range("F6").Value = "Robo y Hurto; salvo que cuente con la Cobertura de Robo y Hurto."
    Range("F7").Value = "Da�os o p�rdidas por explosi�n de Calderas o recipientes a presi�n."
    Range("F8").Value = "Da�o por actos intencionales (culpa grave, dolo o mala fe), negligencia intencional o malevolencia, por el Tomador y/o Asegurado"
    Range("F9").Value = "Retraso, p�rdida de mercado u otros da�os consecuenciales."
    Range("F10").Value = "Las p�rdidas intencionales causadas por el asegurado, sus empleados o representantes."
    
    Range("F12").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
Sub OCEANICA()
    Range("B1").Value = "Equipo Contratista"
    Range("B2").Value = "A: TODO RIESGO MAQUINARIA Y EQUIPO DE CONTRATISTAS."
    Range("B3").Value = "B: RESPONSABILIDAD CIVIL EXTRACONTRACTUAL."
    Range("B4").Value = "C: RESPONSABILIDAD CIVIL OPERACIONES."
    Range("B5").Value = "D: ROBO Y HURTO."

    Range("C1").Value = "DEDUCIBLES"
    Range("C2").Value = "No contratada"
    Range("C3").Value = "No contratada"
    Range("C4").Value = "No contratada"
    Range("C5").Value = "No contratada"
    
    Range("B7").Value = "Condiciones Particulares"
    Range("B8").Value = "Inserte Condiciones Particulares"
    
    Range("B10").Value = "Condiciones Generales"
    Range("B11").Value = "https://1drv.ms/w/s!Au8GQldWcy2ihOlUWQ0ywq5yB1Ic7w?e=gyVUJj"
    
    Range("B13").Value = "Las condiciones particulares pueden variar en las renovaciones, o durante el a�o p�liza por variaciones solicitadas. Las condiciones Generales pueden variar por modificaciones de la aseguradora, pero deben respetar las condiciones pactadas en la vigencia del contrato. Las adjuntas sirven como referencia, puede solicitar las m�s actuales de creerlo necesario."
    
    'Insertando Coberturas generales
    ActiveSheet.Shapes.AddShape(msoShapeCurvedLeftArrow, 19.5, 9, 42.75, 69).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:="", SubAddress:="'Cronograma'!" & lugar
    Range("F1").Value = "PRINCIPALES EXCLUSIONES"
    Range("F2").Value = "Inundaci�n total o parcial por mareas"
    Range("F3").Value = "Robo y Hurto; salvo que cuente con la Cobertura de Robo y Hurto"
    Range("F4").Value = "P�rdidas o Da�os ocasionados por exceder la capacidad de resistencia para lo cual fue dise�ado el bien asegurado; o por ser el bien utilizado para trabajos para los cuales no fue construido."
    Range("F5").Value = "Herramientas, ropa y otros efectos personales u objetos que se encuentren en el equipo."
    Range("F6").Value = "Robo y Hurto; salvo que cuente con la Cobertura de Robo y Hurto."
    Range("F7").Value = "P�rdida o responsabilidades consecuenciales, incluidas: la p�rdida de beneficios, lucro cesante, demora, paralizaci�n del trabajo sea este parcial o totalmente."
    Range("F8").Value = "Multas, impuestos, y/o sanciones impuestas al asegurado "
    Range("F9").Value = "Fallas o desperfectos que exist�an al momento de suscribir este Seguro y que hayan sido del conocimiento del Asegurado o por su direcci�n. "
    Range("F10").Value = "Guerra, invasi�n, actos de una potencia extranjera enemiga, hostilidades u operaciones b�licas "
    
    Range("F13").Value = "La informaci�n suministrada es un resumen, con lo que su asesor considera es lo m�s importante, se recomienda leer las condiciones generales, las cuales son descargables en https://www.sugese.fi.cr/seccion-polizas-registradas/p%C3%B3lizas-vigentes, o las puede solicitar al corredor o a la asistente"
End Sub
