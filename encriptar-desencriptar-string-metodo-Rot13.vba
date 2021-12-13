Function CodRot13(CadenaEnviada As String)
'----------------------------------------------------------------------------------------------------------------------
Fuente            : https://access-global.net/?p=10560
'----------------------------------------------------------------------------------------------------------------------
' Título            : Encriptar cadena de texto
' Autor original    : Desconocido
' Adaptado por      : Ángel Gil
' Actualizado       : 12/03/2002
' Propósito         : Encriptar o desencriptar una cadena de texto enviada como parámetro
' Retorno           : String 
'---------------------------------------------------------------------------------------------------------------------
Function CodRot13(CadenaEnviada As String)
    Dim strAlfabeto As String
    Dim intLongitudCadena As Integer
    Dim intContador As Integer
    Dim strCaracterBuscar As String
    Dim intPosicionCaracter As Integer
    Dim strCadenaSalida As String

    strAlfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    intLongitudCadena = Len(CadenaEnviada)
    For intContador = 1 To intLongitudCadena
        strCaracterBuscar = Mid(CadenaEnviada, intContador, 1)

        '---- Posición que ocupa el caracter dentro del abecedario  ----'

        intPosicionCaracter = InStr(1, strAlfabeto, strCaracterBuscar, 1)

        '----Hay que rotar el caracter 13 veces hacia la izquierda-----'

        If intPosicionCaracter < 14 Then
            intPosicionCaracter = intPosicionCaracter + 13
        Else
            intPosicionCaracter = intPosicionCaracter - 13
        End If
 
        Select Case strCaracterBuscar
            Case "A" To "Z"
                strCadenaSalida = strCadenaSalida & Mid(strAlfabeto, intPosicionCaracter, 1)
            Case "a" To "z"
                strCadenaSalida = strCadenaSalida & LCase(Mid(strAlfabeto, intPosicionCaracter, 1))

                'La eñe y vocales acentuadas así como los caracteres especiales y números no se codifican
                'se dejan “tal cual”
            Case Else
                strCadenaSalida = strCadenaSalida & strCaracterBuscar
        End Select
    Next

    CodRot13 = strCadenaSalida

End Function
'------------------------------------------------------------------------------------------------------------------------

Function DecryptText(Fuente As String) As String
    Dim strDestino As String
    Dim intContador As Integer
    Dim intLongFuente As Integer
    
    strDestino = Fuente
    intLongFuente = Len(Fuente) + 1
    
    For intContador = 1 To Len(strDestino)
        Mid$(strDestino, intLongFuente - intContador, 1) = Chr$((30 + intContador - Asc(Mid$(Fuente, intContador, 1))) And 255)
    Next intContador
    
    DecryptText = strDestino
End Function

'------------------------------------------------------------------------------------------------------------------------

Function EncryptText(Fuente As String) As String
    Dim strDestino As String
    Dim intContador As Integer
    Dim intLongFuente As Integer
    
    strDestino = Fuente
    intLongFuente = Len(Fuente) + 1
    
    For intContador = 1 To Len(strDestino)
        Mid$(strDestino, intContador, 1) = Chr$((30 + intContador - Asc(Mid(Fuente, intLongFuente - intContador, 1))) And 255)
    Next intContador
    
    EncryptText = strDestino
End Function
