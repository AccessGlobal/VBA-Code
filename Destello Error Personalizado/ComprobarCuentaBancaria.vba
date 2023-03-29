Option Compare Database
Option Explicit

Private Enum codError
    codErrorNull = 1000 'Cuando la cuenta es Null
    codErrCta = 1001 'El número de cuenta es incorrecto
    codErrLongCta = 1002 'La longitud de la cuenta es errónea
    codErrDC = 1003 'El dígito de control es erróneo
End Enum

Public Function CompruebaCuentaBancaria(CuentaBancaria As String) As String
'-----------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/tratamiento-de-errores-errores-personalizados
'-----------------------------------------------------------------------------------------------------------------------
' Título            : CompruebaCuentaBancaria
' Autor original    : Luis Viadel
' Creado            : marzo 2016
' Propósito         : verificación de cuenta bancaria mediante 4 comprobaciones, que generan 4 errores personalizados,
'                     utilizando la función 'CVErr'
'                     Los errores personalizados que se crean siguen la siguiente numeración:
'                     1 Es null
'                     2 Contiene sólo números
'                     3 Longitud incorrecta
'                     4 Dígito de control bancario incorrecto
' Argumentos        : La sintaxis de la función consta de un único argumento
'                     Variable          Modo          Descripción
'-----------------------------------------------------------------------------------------------------------------------
'                     CuentaBancaria    Obligatorio   Cuenta bancaria que queremos comprobar
'-----------------------------------------------------------------------------------------------------------------------
' Retorno           : string con la cuenta bancaria bien construída
' Información       : https://support.microsoft.com/en-us/office/cverr-function-d7fd1f1c-3388-4c60-903c-e476865aa467
'-----------------------------------------------------------------------------------------------------------------------
Dim cuenta As String, Banco As String
Dim Sucursal As String, DC As String
Dim NumeroCuenta As String, codPais As String
Dim i As Integer
Dim IBAN As String
Dim coderr As Variant

'Primera comprobación: si es nulo
    If IsNull(CuentaBancaria) Or CuentaBancaria = vbNullString Then
        coderr = CVErr(codErrorNull)
        GoTo LinError
        Exit Function
    End If
    
'Elimina todos los espacios para poder comprobar la longitud
    cuenta = Replace(CuentaBancaria, " ", "")

'Longitud de cuenta 20 caracteres numéricos
'Longitud de IBAN 4 caracteres alfanuméricos
'Segunda comprobación: longitud de cuenta
'comprueba si tiene 24 caracteres, en el caso de tener 20, añade IBAN genérico
    If Len(cuenta) = 20 Then
'Comprueba que todos los elementos de la cadena son números
        For i = 1 To Len(cuenta)
            If Not IsNumeric(Left(cuenta, i)) Then
                coderr = CVErr(codErrCta)
                GoTo LinError
            End If
        Next i
'Añade un IBAN genérico
        cuenta = "ES00" & cuenta
    ElseIf Len(cuenta) <> 24 Then
        coderr = CVErr(codErrLongCta)
        GoTo LinError 'Si no tiene 24 caracteres indica un error
    End If
    
'La cadena cuenta es de 24 caracteres
'Deconstruye la cuenta bancaria
    IBAN = Left(cuenta, 4)
    codPais = Left(cuenta, 2)
    NumeroCuenta = Right(cuenta, 10)
    Banco = Right(Left(cuenta, 8), 4)
    Sucursal = Right(Left(cuenta, 12), 4)
    DC = Left(Right(cuenta, 12), 2)

'Tercera comprobación: dígito de control
    If DigitCalculo(Banco, Sucursal, NumeroCuenta) <> DC Then
        coderr = CVErr(codErrDC)
        GoTo LinError
    End If
'Construimos la primera parte correcta del número de cuenta. Es un número de cuenta válido
    cuenta = Banco & " " & Sucursal & " " & DC & " " & NumeroCuenta

'Cálculo del IBAN, sin tener en cuenta el IBAN recibido
    TempVars!CalculoIBAN = IBANCalculo(codPais, cuenta)
    If IBAN <> TempVars!CalculoIBAN Then MsgBox "El IBAN de su cuenta es " & TempVars!CalculoIBAN, vbInformation + vbOKOnly, "Información sobre la cuenta"

    CompruebaCuentaBancaria = TempVars!CalculoIBAN & " " & Banco & " " & Sucursal & " " & DC & " " & NumeroCuenta
    
    TempVars.RemoveAll
    
    Exit Function

LinError:
    Select Case coderr
        
        Case CVErr(codErrorNull)
            MsgBox "El código de cuenta está vacío. Debe indicar un número de cuenta válido", vbExclamation + vbOKOnly, "Error en nº de cuenta"

        Case CVErr(codErrCta)
            MsgBox "El código de cuenta es incorrecto. Solo puede contener números excepto en el código de país", vbExclamation + vbOKOnly, "Error en nº de cuenta"

        Case CVErr(codErrLongCta)
            MsgBox "La longitud de la cuenta es incorrecta", vbExclamation + vbOKOnly, "Error en nº de cuenta"

        Case CVErr(codErrDC)
            MsgBox "El dígito de control es incorrecto", vbExclamation + vbOKOnly, "Error en nº de cuenta"
    
    End Select

End Function

Public Function IBANCalculo(pais As String, cuenta As String) As String
'-----------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/tratamiento-de-errores-errores-personalizados
'-----------------------------------------------------------------------------------------------------------------------
' Título            : CompruebaCuentaBancaria
' Autor original    : Desconocido
' Adaptaado por     : Luis Viadel
' Fecha             : marzo 2016
' Propósito         : cálculo de los dos números del IBAN de cuenta bancaria que acompañan al código de país
' Argumentos        : La sintaxis de la función consta de un único argumento
'                     Variable          Modo          Descripción
'-----------------------------------------------------------------------------------------------------------------------
'                     pais         Obligatorio   código de dos letras indicativo del país (España ES)
'                     cuenta       Obligatorio   número de cuenta bancaria
'-----------------------------------------------------------------------------------------------------------------------
' Retorno           : string con el IBAN calculado
'-----------------------------------------------------------------------------------------------------------------------
Dim letras As String * 26
Dim Dividendo As Integer
Dim resto As Integer, i As Integer
Dim IBAN As String

' Calcula el valor de las letras, las quita y añade el valor al final
    letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    IBAN = cuenta & CStr(InStr(1, letras, Left(pais, 1)) + 9) & CStr(InStr(1, letras, Right(pais, 1)) + 9) & "00"
        
    For i = 1 To Len(IBAN)
        Dividendo = resto & Mid(IBAN, i, 1)
        resto = Dividendo Mod 97
    Next i
        
    IBANCalculo = pais & Format((98 - resto), "00")

End Function

Function DigitCalculo(ByVal sBank As String, ByVal sSubBank As String, ByVal sAccount As String) As String
'-----------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/tratamiento-de-errores-errores-personalizados
'-----------------------------------------------------------------------------------------------------------------------
' Título            : DigitCalculo
' Autor original    : Desconocido
' Adaptaado por     : Luis Viadel
' Fecha             : marzo 2016
' Propósito         : cálculo de los dos números del dígito de control de una cuenta bancaria
' Argumentos        : La sintaxis de la función consta de un único argumento
'                     Variable          Modo          Descripción
'-----------------------------------------------------------------------------------------------------------------------
'                     sBank          Obligatorio   Código de la entidad bancaria
'                     sSubBank       Obligatorio   Código de la sucursal
'                     sAccount       Obligatorio   número de cuenta bancaria
'-----------------------------------------------------------------------------------------------------------------------
' Retorno           : string con la cuenta bancaria bien construída
'-----------------------------------------------------------------------------------------------------------------------
    TempVars!TempDigit = 0
    TempVars!TempDigit = TempVars!TempDigit + Mid(sBank, 1, 1) * 4
    TempVars!TempDigit = TempVars!TempDigit + Mid(sBank, 2, 1) * 8
    TempVars!TempDigit = TempVars!TempDigit + Mid(sBank, 3, 1) * 5
    TempVars!TempDigit = TempVars!TempDigit + Mid(sBank, 4, 1) * 10
    TempVars!TempDigit = TempVars!TempDigit + Mid(sSubBank, 1, 1) * 9
    TempVars!TempDigit = TempVars!TempDigit + Mid(sSubBank, 2, 1) * 7
    TempVars!TempDigit = TempVars!TempDigit + Mid(sSubBank, 3, 1) * 3
    TempVars!TempDigit = TempVars!TempDigit + Mid(sSubBank, 4, 1) * 6
    TempVars!TempDigit = 11 - (TempVars!TempDigit Mod 11)
    
    If TempVars!TempDigit = 11 Then
        DigitCalculo = "0"
    ElseIf TempVars!TempDigit = 10 Then
        DigitCalculo = "1"
    Else
        DigitCalculo = Format(TempVars!TempDigit, "0")
    End If
    
    TempVars!TempDigit = 0
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 1, 1) * 1
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 2, 1) * 2
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 3, 1) * 4
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 4, 1) * 8
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 5, 1) * 5
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 6, 1) * 10
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 7, 1) * 9
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 8, 1) * 7
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 9, 1) * 3
    TempVars!TempDigit = TempVars!TempDigit + Mid(sAccount, 10, 1) * 6
    TempVars!TempDigit = 11 - (TempVars!TempDigit Mod 11)
    
    If TempVars!TempDigit = 11 Then
        DigitCalculo = DigitCalculo + "0"
    ElseIf TempVars!TempDigit = 10 Then
        DigitCalculo = DigitCalculo + "1"
    Else
        DigitCalculo = DigitCalculo + Format(TempVars!TempDigit, "0")
    End If

    TempVars.RemoveAll
    
End Function

