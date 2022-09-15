Function CalcularDNI(sDNI As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-comprueba-un-dni-o-un-nie
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CalcularDNI
' Autor             : Desconocido
' Fecha             : Desconocida
' Propósito         : Valida un DNI o un NIE y/o calcula la letra del documento
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://www.boe.es/diario_boe/txt.php?id=BOE-A-2008-3580
' Más información   : el cálculo de la letra del DNI se realiza siguiendo los siguientes pasos:
'                     Caso 1: DNI
'                     1. Tomamos el DNI sin la letra y lo dividimos por 23
'                     2. Tomamos la parte entera del valor obtenido en 1
'                     3. Multiplicamos el número del paso 2, de nuevo por 23
'                     4. Al DNI le restamos el número del paso 3
'                     5. Buscamos el número obtenido en el paso 4 en la siguiente relación:
'                        T = 0
'                        R = 1
'                        W = 2
'                        A = 3
'                        G = 4
'                        M = 5
'                        Y = 6
'                        F = 7
'                        P = 8
'                        D = 9
'                        X = 10
'                        B = 11
'                        N = 12
'                        J = 13
'                        Z = 14
'                        S = 15
'                        Q= 16
'                        V = 17
'                        H = 18
'                        L = 19
'                        C = 20
'                        K = 21
'                        E = 22
'                     Caso 2: NIE
'                     1. Tomamos el NIE sin la letra final y lo dividimos por 23
'                     2. Tomamos la parte entera del valor obtenido en 1
'                     3. Multiplicamos el número del paso 2, de nuevo por 23
'                     4. Al DNI le restamos el número del paso 3
'                     5. Buscamos el número obtenido en el paso 4 en la relación anterior.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CalcularDNI_test()
' Dim sDNI As String
'
'    sDNI = "X1234567L" 'NIE
'    sDNI = "12345678Z" 'DNI
'    Debug.Print CalcularDNI(sDNI)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim miNIE As String, miDNI As String
Dim NIEsinletra As String, DNIsinletra As String
Dim mivar As Integer, mivar1 As Integer
Dim DNICompleto As String

    On Error Resume Next

'Eliminamos todos los caracteres de la cadena sDNI, dejando solamente los números y la letra o letras según sea DNI o NIE
    DNICompleto = Replace(sDNI, ".", "")
    DNICompleto = Replace(DNICompleto, "-", "")

'Capturamos el primer caracter para saber si es un DNI o un NIE.
' Si comienza por un número, es un DNI y si comienza por una letra es un NIE
    miNIE = right(DNICompleto, 1)
'Tomamos el valor ASCII del dígito para saber si es número o letra
    mivar1 = Asc(miNIE)

    miDNI = right(DNICompleto, 1)
    mivar = Asc(miDNI)

'Tomamos la numeración sin el último dígito, que siempre será una letra, para poder calcularlo nosotros _
 y validar así la numeración
    NIEsinletra = left(DNICompleto, 8)
    DNIsinletra = left(DNICompleto, 8)

'Si el primer caracter es un número, su valor ASCII estará entre 47 y 58
If mivar1 > 47 And mivar1 < 58 Then

    If mivar > 47 And mivar < 58 Then
    'Si el último carácter no es una letra
        CalcularDNI = sDNI + letra_dni(sDNI)
    Else
    'Si el último carácter es una letra
        If miDNI = letra_dni(DNIsinletra) Then
        'Si el último carácter es una letra y la letra es correcta
            CalcularDNI = sDNI
        Else
            'Si el último carácter es una letra y la letra no es correcta
            CalcularDNI = DNIsinletra + letra_dni(DNIsinletra)
            
            MsgBox "La letra del DNI introducida es errónea. Debería ser " & letra_dni(DNIsinletra), vbInformation, "¡ATENCIÓN! Aviso de MiApp"
        End If
    End If
Else
    If miNIE = letra_dni(NIEsinletra) Then
    'Si el último carácter es una letra y la letra es correcta
        CalcularDNI = sDNI
    Else
    'Si el último carácter es una letra y la letra no es correcta
        CalcularDNI = NIEsinletra + letra_dni(NIEsinletra)
            
        MsgBox "La última letra del NIE introducida es errónea. Debería ser " & letra_dni(DNIsinletra), vbInformation, "¡ATENCIÓN! Aviso de MiApp"
    End If
End If

End Function

Function letra_dni(DNI As String) As String

'-----------------------------------------------------------------------------------------------------------------------------------------------
Select Case left$(DNI, 1) 'Orden EHA/451/2008, de 20 de febrero
    Case Is = "X"
        letra_dni = Mid$("TRWAGMYFPDXBNJZSQVHLCKE", (Val(Replace(DNI, "X", "0")) Mod 23) + 1, 1)
    Case Is = "Y"
        letra_dni = Mid$("TRWAGMYFPDXBNJZSQVHLCKE", (Val(Replace(DNI, "Y", "1")) Mod 23) + 1, 1)
    Case Is = "Z"
        letra_dni = Mid$("TRWAGMYFPDXBNJZSQVHLCKE", (Val(Replace(DNI, "Z", "2")) Mod 23) + 1, 1)
    Case Else
        letra_dni = Mid$("TRWAGMYFPDXBNJZSQVHLCKE", (Val(DNI) Mod 23) + 1, 1)
End Select

End Function


