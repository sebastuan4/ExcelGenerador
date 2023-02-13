Attribute VB_Name = "Main"
Public tipo_poliza  As String
Public aseguradora  As String
Public lugar        As String
Public numero_poliza As String
Sub agregar()
    'Insertando Hoja
    Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = numero_poliza
    'Dejando la hoja en blanco
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub generar()
Attribute generar.VB_ProcData.VB_Invoke_Func = "m\n14"
    tipo_poliza = ActiveCell.Value
    aseguradora = ActiveCell.Offset(0, -1).Value
    numero_poliza = ActiveCell.Offset(0, -2).Value
    lugar = ActiveCell.Offset(0, -3).Address
    Call agregar
    Select Case tipo_poliza
        'Automovil
        Case "AUTOMOVIL"
            Select Case aseguradora
                Case "INS"
                    Call Autos.ins
                Case "LAFISE"
                    Call Autos.Lafise
                Case "QUALITAS"
                    Call Autos.QUALITAS
                Case "OCEÁNICA"
                    Call Autos.OCEANICA
            End Select
        'CREDIAUTO
        Case "CREDIAUTO"
            Select Case aseguradora
                Case "INS"
                    Call Crediauto.ins
            End Select
        'CARGA INTERNACIONAL
        Case "CARGA INTERNACIONAL DOLARES"
            Select Case aseguradora
                Case "OCEÁNICA"
                    Call Carga_Internacional_Dolares.OCEANICA
            End Select
        'CARGA
        Case "CARGA DOLARES"
            Select Case aseguradora
                Case "ASSA"
                    Call Carga_Dolares.ASSA
            End Select
        'Riesgos del trabajo
        Case "RIESGOS DEL TRABAJO"
            Select Case aseguradora
                Case "INS"
                    Call RT.ins
                Case "OCEÁNICA"
            End Select
        'TODO RIESGO INDUSTRIAL Y COMERCIAL COLONES
        Case "TODO RIESGO INDUSTRIAL Y COMERCIAL COLONES"
            Select Case aseguradora
                Case "OCEÁNICA"
                    Call Todo_Riesgo_Comercial_Colones.OCEANICA
            End Select
        'Gastos Medicos Empresarial
        Case "GASTOS MEDICOS"
            Select Case aseguradora
                Case "MAPFRE"
                    Call Gastos_Medicos.Mapfre
                Case "PANAMERICAN"
                    Call Gastos_Medicos.PANAMERICAN
                Case "ASSA"
                    Call Gastos_Medicos.ASSA
            End Select
        'Solo Vida Empresarial
        Case "VIDA EMPRESARIAL"
            Select Case aseguradora
                Case "ASSA"
                    Call Vida_Empresarial.ASSA
            End Select
        'Solo Vida
        Case "VIDA PERSONAL"
            Select Case aseguradora
                Case "INS"
                    Call Vida_Individual.ins
            End Select
        Case "VIDA UNIVERSAL PLUS COLONES"
            Select Case aseguradora
                Case "INS"
                    Call VIDA_PLUS_COLONES.ins
        End Select
        'Incendio Multiriesgo
        Case "INCENDIO MULTIRIESGO"
            Select Case aseguradora
                Case "INS"
                    Call Incendio_MultiRiesgo.ins
                Case "OCEÁNICA"
                    Call IncendioMultiRiesgo.OCEANICA
            End Select
         'Incendio Comercial
        Case "INCENDIO COMERCIAL COLONES"
            Select Case aseguradora
                Case "INS"
                    Call Incendio_Comercial_Colones.ins
            End Select
        'Incendio
        Case "INCENDIO TODO RIESGO COLONES"
            Select Case aseguradora
                Case "LAFISE"
                    Call Incendio_Todo_Riesgo_Colones.Lafise
            End Select
        'Responsabilidad Civil Vigilancia
        Case "RESPONSABILIDAD CIVIL VIGILANCIA"
            Select Case aseguradora
                Case "INS"
                    Call RC_Vigilancia.ins
            End Select
        'Responsabilidad Civil General
        Case "RESPONSABILIDAD CIVIL"
            Select Case aseguradora
                Case "INS"
                    Call RC.ins
                Case "OCEÁNICA"
                    Call RC.OCEANICA
                Case "LAFISE"
                    Call RC.Lafise
                Case "MAPFRE"
                    Call RC.Mapfre
            End Select
        'Responsabilidad Civil Productos
        Case "RESPONSABILIDAD CIVIL Productos"
            Select Case aseguradora
                Case "ASSA"
                    Call RC_Productos.ASSA
            End Select
        'Responsabilidad Civil Profesional
        Case "RESPONSABILIDAD CIVIL Productos"
            Select Case aseguradora
                Case "INS"
                    Call RC_Productos.ASSA
            End Select
        'RESPONSABILIDAD UMBRELLA
        Case "RESPONSABILIDAD UMBRELLA"
            Select Case aseguradora
                Case "OCEÁNICA"
                    Call Responsabilidad_Umbrella.OCEANICA
            End Select
        'Robo Comercial e Industrial
        Case "ROBO COMERCIAL COLONES"
            Select Case aseguradora
                Case "INS"
                    Call Robo_Comercial_Colones.ins
            End Select
        'Equipo Contratista
        Case "EQUIPO DE CONTRATISTA DOLARES"
            Select Case aseguradora
        Case "LAFISE"
                Call Equipo_Contratista_Dolares.Lafise
        Case "OCEÁNICA"
                Call Equipo_Contratista_Dolares.OCEANICA
            End Select
        'Equipo Contratista
        Case "EQUIPO DE CONTRATISTA COLONES"
            Select Case aseguradora
        Case "LAFISE"
                Call Equipo_Contratista.Lafise
            End Select
        'Equipo Electtronico
        Case "EQUIPO ELECTRONICO"
            Select Case aseguradora
        Case "INS"
                Call Equipo_Electronico.ins
            End Select
        'Responsabilidad Civil Productos
        Case "TRANSPORTE MERCANCIAS"
            Select Case aseguradora
        Case "INS"
                Call Transporte_Mercancias.ins
            End Select
         'Hogar comprensivo
        Case "HOGAR"
            Select Case aseguradora
        Case "INS"
                Call Hogar.ins
        Case "OCEÁNICA"
                Call Hogar.OCEANICA
            End Select
         'Hogar comprensivo
        Case "HOGAR 2000"
            Select Case aseguradora
        Case "INS"
                Call Hogar_2000.ins
            End Select
        'Fidelidad
        Case "FIDELIDAD POSICIONES COLONES"
            Select Case aseguradora
        Case "INS"
                Call Fidelidad.ins
                End Select
        'Fidelidad
        Case "FIDELIDAD POSICIONES COLONES"
            Select Case aseguradora
        Case "INS"
                Call Fidelidad.ins
                End Select
        'VALORES EN TRÁNSITO COLONES
        Case "VALORES EN TRÁNSITO COLONES"
            Select Case aseguradora
        Case "INS"
                Call VALORES_Transito.ins
                End Select
         'MÉDICO COLECTIVO
        Case "MÉDICO COLECTIVO"
            Select Case aseguradora
        Case "INS"
                Call Medico_Colectivo.ins
                End Select
        Case Else
            MsgBox "No pudimos encontrar este tipo de poliza, por favor comunicarnos " & tipo_poliza & " " & aseguradora
    End Select
Call estica.dar_estetica
End Sub

