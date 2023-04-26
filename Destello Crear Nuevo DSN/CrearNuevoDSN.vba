Option Compare Database
Option Explicit

Public Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As Long, ByVal frequest As Long, ByVal lpszdriver As String, ByVal lpszattributes As String) As Long

Public Function CrearNuevoDSN(nombreDSN As String) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-modificar-o-eliminar-origen-de-datos-en-tiempo-de-ejecucion/
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CrearNuevoDSN
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : 12/05/2013
' Propósito         : crear o modificar un origen de datos remoto
' Argumentos        : ninguno
'---------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/sql/odbc/reference/syntax/sqlconfigdatasource-function?view=sql-server-ver16
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub CrearNuevoDSN_test()
' Dim resultado as long
'
'      resultado=CrearnuevoDSN("MiConexion")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim atts As String
Dim strDriver As String

       
    strDriver = "postgresql Unicode"
    atts = "DSN=" & nombreDSN & Chr(0) & _
            "DATABASE=" & "MiBD" & Chr(0) & _
            "SERVER=" & "MiServidor" & Chr(0) & _
            "PORT=" & "5432" & Chr(0) & _
            "UID=" & "MiUsuario" & Chr(0) & _
            "PWD=" & "MiPass" & Chr(0)
            
    CrearNuevoDSN = SQLConfigDataSource(0, 1, strDriver, atts)

End Function


