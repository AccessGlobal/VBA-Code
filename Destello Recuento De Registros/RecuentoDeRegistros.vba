Function RecuentoDeRegistros() As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-contar-registros
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : RecuentoDeRegistros
' Autor             : Luis Viadel | luisviadel@cowtechnologies.net | https://cowtechnologies.net
' Creación          : octubre 22
' Propósito         : ver tres diferentes formas de contar registros
' Retorno           : el número de registros de la tabla o consulta
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/office/vba/access/concepts/data-access-objects/count-the-number-of-records-in-a-dao-recordset
'                     https://learn.microsoft.com/es-es/office/vba/api/access.application.dcount?redirectedfrom=MSDN
'                     https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/recordcount-property-ado?view=sql-server-ver16
'                     Microsoft ActiveX Data Objects 2.8 Library
'-----------------------------------------------------------------------------------------------------------------------------------------------

'Método 1: DAO recordset
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM facvta")
        rstTable.MoveLast
            RecuentoDeRegistros = rstTable.RecordCount
            Debug.Print "Registros por el método 1: " & RecuentoDeRegistros
        rstTable.Close
    Set rstTable = Nothing
    
'Método 2: ADO recordset
Dim strSql As String
Dim objCnn As ADODB.Connection
Dim objRst As ADODB.Recordset
    
    strSql = "select * from facvta"
    
    Set objCnn = CurrentProject.Connection
        Set objRst = New ADODB.Recordset
            objRst.CursorLocation = adUseClient
'            Set objRst = objCnn.Execute(strSql)
            objRst.Open strSql, CurrentProject.Connection
            
            RecuentoDeRegistros = objRst.RecordCount
            Debug.Print "Registros por el método 2: " & RecuentoDeRegistros
            objRst.Close
        Set objRst = Nothing
    Set objCnn = Nothing
    

'Método 3: DCount
    RecuentoDeRegistros = DCount("[idfacvta]", "facvta")
    Debug.Print "Registros por el método 3: " & RecuentoDeRegistros

End Function
