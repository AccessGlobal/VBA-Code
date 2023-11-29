Public Function bcheastrBuscaAcentos(X As String) As String
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-buscar-acentos/
'                     Destello formativo 385
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : bcheastrBuscaAcentos
' Autor original    : J.Bengoechea
' Creado            : 19/03/2001
' Actualizado       : McPegasus | rafaelandrada@access-global.net
'                   : 19/06/2003
' Propósito:        : Convertir una vocal en una cadena que contiene la misma vocal con todas las formas de acentuación posibles.
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     X              Obligatorio      Cadena que queremos convertir
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub bcheastrBuscaAcentos_test()
' Dim cadenabuscar as string
'
'      cadenabuscar = bcheastrBuscaAcentos (MiCadena)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Letra As String
Dim Vocal As String
Dim Nuevaletra As String

Dim I As Variant
Dim A As Integer
Dim L As Integer
Dim busc As String
    
Static letras(6) As Variant
    
    L = Len(X)
    busc = X
    A = 1

    letras(1) = "AÁÀÂÄ"
    letras(2) = "EÉÈÊË"
    letras(3) = "IÍÌÎÏ"
    letras(4) = "OÓÒÔÖ"
    letras(5) = "UÚÙÛÜ"
'Fecha:         19/06/2003
'Desarrollador: McPegasus.
'Modificación:  Antes "YÝýÿ", al tener distinta cantidad de carácteres que el resto de letras, se produce un error.
    letras(6) = "yYÝýÿ"
    
    While A <= L
        Letra = Mid(busc, A, 1)
        For Each I In letras
            Vocal = InStr(1, I, Letra, 1)
            If Vocal > 0 Then
                Nuevaletra = "[" & I & "]"
                busc = Left(busc, A - 1) & Nuevaletra & Right(busc, L - A)
                A = A + 1 + Len(I)
                L = L + 1 + Len(I)
                Exit For
            End If
        Next
        A = A + 1
    Wend

    If busc = "" Then
        bcheastrBuscaAcentos = X
    Else
        bcheastrBuscaAcentos = busc
    End If

End Function
