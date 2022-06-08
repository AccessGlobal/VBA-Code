Public Function GetMyPublicIP() As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-obtener-mi-ip-publica
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetMyPublicIP
' Autor original    : Christos Samaras | xristos.samaras@gmail.com
' Fuente original   : http://www.myengineeringworld.net
' Adaptado          : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : 22/11/2014
' Propósito         : obtener mi IP pública de conexión a Internet
' Referencia        : Microsoft XML, v6.0
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms757849(v=vs.85)?redirectedfrom=MSDN
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub OrdenaForm_test()
'
'   Debug.print GetMyPublicIP
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim HttpRequest As Object
    
    On Error Resume Next
    
    Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
        If Err.Number <> 0 Then
            GetMyPublicIP = "No se puede crear el objeto XMLHttpRequest"
            Set HttpRequest = Nothing
            Exit Function
        End If
    On Error GoTo 0
        
    HttpRequest.Open "GET", "http://myip.dnsomatic.com", False
        
    HttpRequest.Send
            
    GetMyPublicIP = HttpRequest.responseText
    
End Function