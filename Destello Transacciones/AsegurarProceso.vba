Public Function AsegurarProceso() As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-transacciones
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AsegurarProceso
' Autor original    : Luis Viadel
' Creado            : marzo 2010
' Adaptado          : diciembre 2022
' Propósito         : asegurar cualquier proceso mediante el uso de transacciones
' Retorno           : verdadero / falso según finalice el proceso satisfactoriamente o no
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     En este caso no hemos utilizado argumentos. Podríamos crear una función a la que le pasásemos el proceso concreto
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/es-es/office/vba/access/concepts/data-access-objects/use-transactions-in-a-dao-recordset
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
'
'   If AsegurarProceso then
'       Realizamos operaciones
'   Else
'       Mandamos un mensaje de error
'   End If
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Connect As ADODB.Connection
Dim rstTable As ADODB.Recordset
Dim strConnect As String
    
    On Error GoTo LinErr
    
    Set Connect = New ADODB.Connection
    
        Connect.Mode = adModeShareExclusive
        Connect.IsolationLevel = adXactIsolated
        
        Connect.Open "MyDSN", "MyDB", "MyPassword"

        Set rstTable = New ADODB.Recordset
            rstTable.Open "MyTable", Connect, adOpenDynamic, adLockPessimistic, adCmdTable
    
'Bloque de transacción. Aquí realizamos nuestras operaciones
        Connect.BeginTrans
   
'Operaciones

        Connect.CommitTrans
        
'Cierre de la conexión y los objetos utilizados
            rstTable.Close
        Connect.Close
        Set rstTable = Nothing
    Set Connect = Nothing
    
    AsegurarProceso = True
    
    Exit Function
    
LinErr:
'Deshacemos la transacción
    Connect.RollbackTrans
    
    AsegurarProceso = False
    
    MsgBox "Se ha producido el error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description

End Function
