Public Function GetMyLocalIP()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-obtener-mi-ip-local
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetMyLocalIP
' Autor             : desconocido
' Adaptado          : Luis Viadel | https://cowtechnologies.net
' Propósito         : obtener mi IP en la red local en la que me encuentro
' Retorno           : devuelve la dirección IP de mi equipo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemservices-execquery
'                     https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemobjectset
'                     https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration
'                     https://docs.microsoft.com/en-us/windows/win32/wmisdk/querying-with-wql
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y abrir el
'                     formulario para ver su funcionamiento.
'
'Sub GetMyLocalIP_test()
'
'   Call GetMyLocalIP
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim strComputer     As String
Dim objWMIService   As Object
Dim colItems        As Object
Dim objItem         As Object
Dim myIPAddress     As String
Dim strQuery        As String

    strComputer = "."
       
'Método 1

    strQuery = "SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery(strQuery)
            For Each objItem In colItems
                If Not IsNull(objItem.IPAddress) Then myIPAddress = Trim(objItem.IPAddress(0))
                Debug.Print "Método 1: " & myIPAddress
            Next
        Set colItems = Nothing
    Set objWMIService = Nothing

'Método 2
    strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"

    Set objWMIService = GetObject("winmgmts://./root/CIMV2")
        Set colItems = objWMIService.ExecQuery(strQuery, "WQL", 16)
            For Each objItem In colItems
                If IsArray(objItem.IPAddress) Then
                    myIPAddress = Join(objItem.IPAddress, " | ")
                    Debug.Print "Método 2: " & myIPAddress
                End If
            Next
        Set colItems = Nothing
    Set objWMIService = Nothing

End Function
