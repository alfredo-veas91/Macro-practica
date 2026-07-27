Attribute VB_Name = "Módulo1"
Function DiferenciaFechasEnMeses(FechaDada As Date) As Integer
    Dim FinDeAño As Date
    
    ' Establecer la fecha del 31 de diciembre del mismo año
    FinDeAño = DateSerial((Year(FechaDada)), 12, 31)
    
    ' Calcula la diferencia en meses
    DiferenciaFechasEnMeses = DateDiff("m", FechaDada, FinDeAño)
End Function

Function DiferenciaInicioAño(FechaAux As Date) As Integer
    'Dim FechaDada As Date
    Dim InicioAño  As Date
    Dim Diferencia As Integer
    
    ' Lee la fecha desde una celda específica, por ejemplo, C8
    On Error Resume Next
    FechaDada = Range("C8").Value
    On Error GoTo 0
    
    ' Verificar si se pudo leer una fecha válida
    If FechaDada = 0 Then
        MsgBox "La celda no contiene una fecha válida.", vbExclamation
        'Exit Sub
    End If
    
    ' Establecer la fecha del 31 de diciembre del mismo año
    InicioAño = DateSerial((Year(FechaDada)), 1, 1)
    
    ' Calcula la diferencia en vbMonday
    DiferenciaInicioAño = DateDiff("m", InicioAño, FechaAux)
    
    ' Puedes hacer algo con el valor de Diferencia aquí, como mostrarlo en un MsgBox
    'MsgBox "La diferencia en meses es: " & Diferencia
End Function


Sub Macro()
    Dim FechaDada As Date
    Dim FinDeAño As Date
    Dim FechaAux As Date
    Dim Fila As Integer
    Dim DiferenciaMeses As Integer
    Dim Anio As Integer
    Dim ValorInic As Double
    Dim Incorp As Double
    Dim Erog As Double
    Dim Desmantela As Double
    Dim Desarrollo As Double
    Dim Gastos As Double
    Dim Baja As Double
    Dim ValorBien As Double
    Dim Deprec As Double
    Dim Deter As Double
    Dim VidaUAnios As Double
    Dim VidaUMeses As Double
    Dim MesesDeprec As Double
    Dim DeprecPeriod As Double
    Dim ValorDeprec As Double
    
   
    ' Leer la fecha
    On Error Resume Next
    FechaDada = Range("C8").Value
    On Error GoTo 0
    
    ' Verificar si se pudo leer una fecha válida
    If FechaDada = 0 Then
        MsgBox "La celda no contiene una fecha válida.", vbExclamation
        Exit Sub
    End If
    
    FechaAux = FechaDada
    
    ' Calcula meses a depreciar
    MesesDeprec = DiferenciaFechasEnMeses(FechaDada)
    
    ' Sacar el año
    Anio = Year(FechaDada)
       
    Fila = 13
    
    ' Inicalizar valores
    ValorInic = Cells(Fila, "D").Value
    Incorp = Cells(Fila, "E").Value
    Erog = Cells(Fila, "F").Value
    Desmantela = Cells(Fila, "G")
    Desarrollo = Cells(Fila, "H")
    Gastos = Cells(Fila, "I")
    Baja = Cells(Fila, "J")
    ValorBien = ValorInic + Incorp + Erog + Desmantela + Desarrollo + Gastos + Baja
    Cells(Fila, "K").Value = ValorBien
    Deperc = Cells(Fila, "L").Value
    Deter = Cells(Fila, "M").Value
    VidaUAnios = Cells(Fila, "N").Value
   
    VidaUMeses = VidaUAnios * 12
    Cells(Fila, "O").Value = VidaUMeses
    Cells(Fila, "P").Value = MesesDeprec
    
    ' Hacer los calculos
    
    
    Cells(Fila, "B").Value = Anio
    DeprecPeriod = ((ValorBien - Deprec - Deter) / VidaUMeses) * MesesDeprec
    DeprecPeriod = Round(DeprecPeriod, 0)
    Cells(Fila, "Q").Value = DeprecPeriod
    ValorDeprec = ValorBien - Deperc - Deter - DeprecPeriod
    Cells(Fila, "R").Value = ValorDeprec
        ' Cambio de fila
    Do
        Fila = Fila + 1
        VidaUAnios = VidaUAnios - 1
        Cells(Fila, "D").Value = ValorInic
        ' Separador
        Cells(Fila, "N").Value = VidaUAnios
        Deprec = Deprec + DeprecPeriod
        Cells(Fila, "L").Value = Deprec
        VidaUMeses = Cells(Fila - 1, "O") - Cells(Fila - 1, "P")
        Cells(Fila, "O").Value = VidaUMeses
        Incorp = Cells(Fila, "E").Value
        Erog = Cells(Fila, "F").Value
        Desmantela = Cells(Fila, "G")
        Desarrollo = Cells(Fila, "H")
        Gastos = Cells(Fila, "I")
        Baja = Cells(Fila, "J")
        ValorBien = ValorInic + Incorp + Erog + Desmantela + Desarrollo + Gastos + Baja
        Cells(Fila, "K").Value = ValorBien
        Anio = Anio + 1
        Cells(Fila, "B") = Anio
        If VidaUAnios > 0 Then
            MesesDeprec = 12
        ElseIf VidaUAnios = 0 Then
            MesesDeprec = DiferenciaInicioAño(FechaAux) + 1
        End If
        Cells(Fila, "P").Value = MesesDeprec
        If VidaUAnios = 0 Then
           DeprecPeriod = (((ValorBien - Deprec - Deter) / VidaUMeses) * MesesDeprec) - 1
        Else
           DeprecPeriod = ((ValorBien - Deprec - Deter) / VidaUMeses) * MesesDeprec
        End If
        DeprecPeriod = Round(DeprecPeriod, 0)
        Cells(Fila, "Q").Value = DeprecPeriod
        ValorDeprec = Cells(Fila - 1, "R").Value - DeprecPeriod
        Cells(Fila, "R").Value = ValorDeprec
    Loop While VidaUAnios > 0
        
End Sub

