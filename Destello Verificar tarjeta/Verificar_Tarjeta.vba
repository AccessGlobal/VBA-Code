Option Compare Database
Option Explicit

Public Function VerificaTarjeta(numtarjeta As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-verificar-tarjeta-de-credito/
'                     Destello formativo 338
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : VerificaTarjeta
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 2019
' Propósito         : verificar la numeración de una tarjeta de crédito, la fecha de caducidad y el dígito de control
' Argumento         : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     numTarjeta     Obligatorio      numero de la tarjeta que se desea verificar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://www.validcreditcardnumber.com/
'                     https://en.wikipedia.org/wiki/Luhn_algorithm
' Más información   : Estructura de la numeración
'                     Dígito 1      : empresa emisora
'                     Dígito 2      : tipo de tarjeta (crédito, débito)
'                     Dígitos 3 a 7 : número de cuenta
'                     Dígitos 8 a 14: número de tarjeta. Permite asociar la tarjeta a la cuenta del usuario
'                     Dígito 15     : número de control. Debe cumplir con el algoritmo de Luhn

'                     Dígitos en una tarjeta:
'                       Visa y Visa Electron: 13 o 16
'                       Mastercard: 16
'                       American Express: 15
'                       Diner 's Club: 14 ( incluyendo enRoute, International, Blanche )
'                       Maestro: 12 a 19 ( tarjeta de débito multinacional )
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Private Sub btnQRTest_Click()
'Dim resultado As boolean
'
'       resultado= VerificaTarjeta (Me.numtarjeta)
'
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim str As String
Dim lng As Long
Dim mesTarjeta As Long, anoTarjeta As Long
Dim resultadoLuhn As Integer

    str = Left(numtarjeta, 1)
    lng = Len(numtarjeta)
    
'Número inicial
'La primera condición es que el número inicial de las tarjetas de crédito no puede ser ninguno de estos dígitos 1,2,7,8,9,0
    If InStr(1, "127890", str) <> 0 Then
        GoTo Linfalse
    End If
'Si se trata de una American Express tendrá 15 caracterres y comenzará por 3 o por 4
    If lng = 15 Then 'Es una tarjeta American Express que debe cumplir otra condición
        If str <> 3 Or str <> 4 Then
            GoTo Linfalse
        End If
    ElseIf lng <> 16 Then GoTo Linfalse
    
    End If
'Verifica la fecha de caducidad
    mesTarjeta = Form_Test.mescaducidad
    anoTarjeta = "20" & Form_Test.anoCaducidad
       
    If anoTarjeta < Year(Date) Or anoTarjeta > Year(Date) + 10 Then
        GoTo Linfalse
    ElseIf anoTarjeta = Year(Date) And mesTarjeta < Month(Date) Then
        GoTo Linfalse
    End If
'Verifica el algoritmo de Luhn
    If luhnCheckSum(numtarjeta) = 0 Then
        VerificaTarjeta = True
        Exit Function
    End If
         
Linfalse:
    VerificaTarjeta = False
 
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-verificar-tarjeta-de-credito/
'                     Destello formativo 338
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : luhnSum
' Autor             : Christian Mäder | https://gist.github.com/cimnine/7913819.js
' Creado            : 2021
' Propósito         : verificar el último dígito de una tarjeta de crédito
'-----------------------------------------------------------------------------------------------------------------------------------------------
Function luhnSum(InVal As String) As Integer
Dim evenSum As Integer
Dim oddSum As Integer
Dim strLen As Integer
         
    evenSum = 0
    oddSum = 0
     
    strLen = Len(InVal)
     
    Dim i As Integer
    For i = strLen To 1 Step -1
        Dim digit As Integer
        digit = CInt(Mid(InVal, i, 1))
         
        If ((i Mod 2) = 0) Then
            oddSum = oddSum + digit
        Else
            digit = digit * 2
             
            If (digit > 9) Then
                digit = digit - 9
            End If
             
            evenSum = evenSum + digit
        End If
    Next i
     
    luhnSum = (oddSum + evenSum)
    
End Function

' for the curious
Function luhnCheckSum(InVal As String)
    
    luhnCheckSum = luhnSum(InVal) Mod 10

End Function

' true/false check
Function luhnCheck(InVal As String) As Integer
    
    luhnCheck = (luhnSum(InVal) Mod 10) = 0

End Function

' returns a number which, appended to the InVal, turns the composed number into a valid luhn number
Function luhnNext(InVal As String)
Dim luhnCheckSumRes
    
    luhnCheckSumRes = luhnCheckSum(InVal)
     
    If (luhnCheckSumRes = 0) Then
        luhnNext = 0
    Else
        luhnNext = ((10 - luhnCheckSumRes))
    End If

End Function
