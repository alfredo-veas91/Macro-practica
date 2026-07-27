Attribute VB_Name = "Módulo3"
Sub DiferenciaInicioAño()
    'Dim FechaDada As Date
    Dim InicioAño  As Date
    Dim Diferencia As Integer
    
    ' Lee la fecha desde una celda específica, por ejemplo, C8
    On Error Resume Next
    FechaDada = Range("E26").Value
    On Error GoTo 0
    
    ' Verificar si se pudo leer una fecha válida
    If FechaDada = 0 Then
        MsgBox "La celda no contiene una fecha válida.", vbExclamation
        'Exit Sub
    End If
    
    ' Establecer la fecha del 31 de diciembre del mismo año
    InicioAño = DateSerial((Year(FechaDada)), 1, 1)
    
    ' Calcula la diferencia en vbMonday
    Diferencia = DateDiff("m", InicioAño, FechaDada, vbMonday, vbFirstFullWeek)
    
    ' Puedes hacer algo con el valor de Diferencia aquí, como mostrarlo en un MsgBox
    MsgBox "La diferencia en meses es: " & Diferencia
End Sub
