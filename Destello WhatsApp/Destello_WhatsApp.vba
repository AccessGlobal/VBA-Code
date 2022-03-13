Public Function EnvioWhatsapp(ByVal telnumber, ByVal msgtext) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-sencilla-forma-de-enviar-mensaje-de-whatsapp
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : EnvioWhatsapp
' Autor original    : Luis Viadel
' Propósito         : enviar un mensaje a través de WhatsApp
' Retorno           : texto con información del mensaje
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     telnumber     Obligatorio    Número de teléfono móvil
'                     msgtext       Obligatorio    Texto del mensaje
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://www.whatsapp.com/business/api?lang=es
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub test_sendWhatsapp()
'Dim str As String
'
'Debug.Print str = EnvioWhatsapp("NUMERO_TELEFONO", "Mensaje enviado desde Access")
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

CreateObject("Shell.Application").Open "https://wa.me/" + telnumber + "/?text=" & msgtext

Sleep 3000

SendKeys "{ENTER}"

End Function

