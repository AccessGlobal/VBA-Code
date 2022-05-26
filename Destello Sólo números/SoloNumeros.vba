Private Sub txtPrueba_KeyPress(KeyAscii As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-estas-obligado-a-escribir-solo-numeros
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : KeyPress
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Propósito         : obligar al usuario a escribir sólo números en un textbox
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     KeyAscii      Obligatorio    pulsación del teclado
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://thecodeforyou.blogspot.com/2013/01/vb-keyascii-values.html
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código a tu aplicación, coloca este bloque en el evento "Al pulsar una tecla" de cualquier textbox
'-----------------------------------------------------------------------------------------------------------------------------------------------

    If InStr("0123456789", Chr(KeyAscii)) = 0 Then 'Podemos indicar los caracteres que deseamos permitir
        
        If KeyAscii <> 8 Then 'Se permite a la tecla con valor Ascii=8, el retroceso
            KeyAscii = 0
        Else
            Exit Sub
        End If
        
        KeyAscii = 0
    
    End If

End Sub