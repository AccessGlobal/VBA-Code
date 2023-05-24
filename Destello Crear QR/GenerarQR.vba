Public Function QRGenerador(convertirQR As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-la-forma-mas-sencilla-de-crear-un-codigo-qr/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : QRGenerador
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : 10/2018
' Propósito         : obtener código QR en formato de imagen
' Argumento         : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     convertirQR   Obligatorio      cadena de texto que queremos convertir en QR
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://developers.google.com/chart?hl=es-419
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Private Sub btnQRTest_Click()
'
'   Me.QRImage.Picture = QRGenerador(Me.texto)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ruta As String
Dim URLQR As String

    On Error GoTo LinError
    
    ruta = Application.CurrentProject.Path & "\Temp\qrTemp.png"
    
    URLQR = "https://chart.googleapis.com/chart?chs=180x180&cht=qr&chl=" & convertirQR & "&choe=UTF-8"

'Limpia la caché para que pueda descargar el nuevo fichero
    DeleteUrlCacheEntry URLQR
'Descarga el fichero en la ruta indicada
    URLDownloadToFile 0, URLQR, ruta, 0, 0

    QRGenerador = ruta

    Exit Function
    
LinError:
    MsgBox "Se ha producido un error"
    QRGenerador = vbNullString
    
End Function
