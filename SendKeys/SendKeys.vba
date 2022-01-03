Public Sub mcSendKeys(ByVal strKey As String, Optional ByVal blnWait As Boolean)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/sendkeys/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcSendKeys
' Autor original    : De varias ideas y aportaciones de otros colegas de Access User Groups España.
' Adaptado por      : Rafael Andrada .:McPegasus:. de BeeSoftware.
' Actualizado       : 05/02/2021
' Propósito         : Se utiliza la instrucción SendKey para enviar una o más pulsaciones de teclas a la ventana activa como si se escribieran en el teclado. ¡IMPORTANTE! Soluciona el problema de utilizar directamente SendKey que modifica el estado de la tecla BloqNum o NumLock (bloqueo numérico) que lo activa o desactiva.
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strKey            Obligatorio     El valor String especifica las pulsaciones de teclas que enviar.
'                     blnWait           Opcional        El valor Boolean especifica el modo de espera. Si es False (predeterminado), se devuelve el control al procedimiento inmediatamente después de enviar las teclas. Si es True, las pulsaciones de teclas se deben procesar antes de que se devuelva el control al procedimiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/es-es/office/vba/language/reference/user-interface-help/sendkeys-statement?f1url=%3FappId%3DDev11IDEF1%26l%3Des-ES%26k%3Dk(vblr6.chm1009015);k(TargetFrameworkMoniker-Office.Version%3Dv16)%26rd%3Dtrue
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub TITULO_test()
'    Call mcSendKeys("F4", False)            'Si se ejecuta desde el editor de VBA, se abre la ventana Propiedades.
'    Call mcSendKeys("+F2", False)           'Si se coloca en el evento DblClick de un control de un cuado de texto se abre el Zoom.
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'End Function
  
    Dim WshShell                                   As Object
     
    Dim strWork                                    As String
     
    
    Set WshShell = CreateObject("WScript.Shell")
     
    strWork = Left(strKey, 1)
     
    Select Case strWork
        Case "+", "^", "%"
            strKey = Mid(strKey, 2)
            WshShell.SendKeys strWork & "{" & strKey & "}", blnWait
        
        Case Else
            WshShell.SendKeys "{" & strKey & "}", blnWait
     
    End Select
     
    Set WshShell = Nothing

End Sub
