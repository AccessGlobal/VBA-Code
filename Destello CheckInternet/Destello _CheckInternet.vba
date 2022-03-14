'Declarar a nivel de módulo
Public Declare PtrSafe Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Public Const FLAG_ICC_FORCE_CONNECTION = &H1

Public Function CheckInternet(strURL As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-tengo-conexion-a-internet
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CheckInternet
' Autor original    : Luis Viadel
' Creado            : abril 2010
' Propósito         : comprobar la conexión a internet en un momento determinado
' Retorno           : verdadero / falso según tengamos conexión o no
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strURL          Obligatorio      El valor Boolean especifica si tenemos o no conexión
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetcheckconnectiona
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
'
'    If CheckInternet("https://access-global.net") then
'        Debug.Print "Tengo conexión"
'    Else
'        Debug.Print "No tengo conexión"
'    End If
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

CheckInternet = InternetCheckConnection(strURL, FLAG_ICC_FORCE_CONNECTION, 0&)

End Function
