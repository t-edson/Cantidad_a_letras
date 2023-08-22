Public Class Form1
    Private Function Ordinal1(num As String) As String
        'Devuelve el ordinal de un número dígito.
        Select Case num
            Case "1" : Return "uno "
            Case "2" : Return "dos "
            Case "3" : Return "tres "
            Case "4" : Return "cuatro "
            Case "5" : Return "cinco "
            Case "6" : Return "seis "
            Case "7" : Return "siete "
            Case "8" : Return "ocho "
            Case "9" : Return "nueve "
            Case "0" : Return ""
        End Select
    End Function
    Private Function Ordinal10(num As String) As String
        'Devuelve el ordinal de un número de dos o un dígito.
        num = Val(num)  'Quita ceros a la izquierda
        If Len(num) = 1 Then Return Ordinal1(num)
        Dim dig1 As String = Mid(num, 1, 1)
        Dim dig2 As String = Mid(num, 2, 1)
        Select Case num
            Case "10" : Return "diez "
            Case "11" : Return "once "
            Case "12" : Return "doce "
            Case "13" : Return "trece "
            Case "14" : Return "catorce "
            Case "15" : Return "quince "
            Case "20" : Return "veinte "
            Case "30" : Return "treinta "
            Case "40" : Return "cuarenta "
            Case "50" : Return "cincuenta "
            Case "60" : Return "sesenta "
            Case "70" : Return "setenta "
            Case "80" : Return "ochenta "
            Case "90" : Return "noventa "
        End Select
        Select Case dig1
            Case "1" : Return "dieci" & Ordinal1(dig2)
            Case "2" : Return "veinti" & Ordinal1(dig2)
            Case "3" : Return "treinta y " & Ordinal1(dig2)
            Case "4" : Return "cuarenta y " & Ordinal1(dig2)
            Case "5" : Return "cincuenta y " & Ordinal1(dig2)
            Case "6" : Return "sesenta y " & Ordinal1(dig2)
            Case "7" : Return "setenta y " & Ordinal1(dig2)
            Case "8" : Return "ochenta y " & Ordinal1(dig2)
            Case "9" : Return "noventa y " & Ordinal1(dig2)
        End Select
    End Function
    Private Function ordinal100(num As String) As String
        'Devuelve el ordinal de un número de tres dígitos
        num = Val(num)  'Quita ceros a la izquierda
        If Len(num) < 3 Then Return Ordinal10(num)
        Dim dig1 As String = Mid(num, 1, 1)
        Dim dig23 As String = Mid(num, 2, 2)
        If num = "100" Then
            Return "cien "
        End If
        Select Case dig1
            Case "0" : Return Ordinal10(dig23)
            Case "1" : Return "ciento " & Ordinal10(dig23)
            Case "2" : Return "doscientos " & Ordinal10(dig23)
            Case "3" : Return "trescientos " & Ordinal10(dig23)
            Case "4" : Return "cuatrocientos " & Ordinal10(dig23)
            Case "5" : Return "quinientos " & Ordinal10(dig23)
            Case "6" : Return "seiscientos " & Ordinal10(dig23)
            Case "7" : Return "setecientos " & Ordinal10(dig23)
            Case "8" : Return "ochocientos " & Ordinal10(dig23)
            Case "9" : Return "novecientos " & Ordinal10(dig23)
        End Select
    End Function
    Private Function Extraer3dig(ByRef num As String) As String
        'Devuelve el ordinal de los últimos 3 dígitos extraidos de la cadena "num".
        num = Val(num)  'Quita ceros a la izquierda
        Select Case Len(num)
            Case 0 : Return ""
            Case 1
                Dim tmp As String = num
                num = ""
                Return Ordinal1(tmp)
            Case 2
                Dim tmp As String = num
                num = ""
                Return Ordinal10(tmp)
            Case 3
                Dim tmp As String = num
                num = ""
                Return ordinal100(tmp)
            Case > 3
                Dim tmp As Integer = Mid(num, Len(num) - 2, 3)    'Copia los últimos 3 dígitos como número.
                num = Mid(num, 1, Len(num) - 3)    'Extrae lo que queda
                Return ordinal100(tmp)
        End Select
    End Function
    Public Function CantidadALetras(numero As Double) As String
        'Convierte una cantidad en números a su formato en letras.
        '********************************************************************************
        Dim numStr As String
        Dim pEntera As String
        Dim pDecimal As String
        Dim ordinal As String = ""
        Dim ordinalDec As String = ""
        Dim signo As String = ""
        'Validación
        If numero = 0 Then
            Return "cero con 00/100"
        End If
        'Validación de números negativos
        If numero < 0 Then
            numero = -numero    'Convierte a positivo
            signo = "menos "
        End If
        numStr = Format(numero, "0.00")     'Formatea a 2 decimales (Siempre)
        'Divide parte entera y decimal
        pEntera = numStr.Split(".")(0)
        pDecimal = numStr.Split(".")(1)
        ordinalDec = "con " & pDecimal & "/100"
        '**********proceso de conversión***********
        ordinal &= Extraer3dig(pEntera) 'Lee hasta 999
        If pEntera = "" Then    'No hay miles
            Return signo & ordinal & ordinalDec
        End If
        'Hay bloque de miles
        Dim miles As String = Extraer3dig(pEntera) 'Lee bloque de miles
        If miles <> "" Then 'Hay miles
            If miles = "uno " Then miles = "mil " Else miles &= "mil "
            ordinal = miles & ordinal
        End If
        If pEntera = "" Then    'No hay milllones
            Return signo & ordinal & ordinalDec
        End If
        'Hay bloque de millones
        Dim mills As String = Extraer3dig(pEntera) 'Lee bloque de miles
        If mills <> "" Then
            If mills = "uno " Then mills = "un millón " Else mills = mills & "millones "
            ordinal = mills & ordinal
        End If
        If pEntera = "" Then    'No hay milllones
            Return signo & ordinal & ordinalDec
        End If
        'Hay bloque de miles de millones
        MsgBox("Número excede a 999 millones.", vbExclamation)
        Return ""
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = CantidadALetras(NumericUpDown1.Value)
    End Sub
End Class
