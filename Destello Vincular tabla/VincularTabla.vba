Sub mcVincular_mi_primera_tabla()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-vincular-mi-primera-tabla-de-access/
'                     Destello formativo 277
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcVincular_mi_primera_tabla
' Autor original    : Rafael Andrada | McPegasus |https://beesoftware.es
' Creado            : desconocido
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : vincular una tabla mediante VBA
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-object-dao
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dbsFront                                    As DAO.Database
Dim tdfFront                                    As DAO.TableDef
Dim strBack                                     As String
Dim strTableName                                As String

    Set dbsFront = CurrentDb
    
    strBack = Application.CurrentProject.Path & "\Backd_Test.accdb"
    
    strTableName = "Test"
        
    'Establecer el objeto TableDef (tabla base o tabla adjunta). Es el nombre de la tabla donde están las tablas vinculadas.
    Set tdfFront = dbsFront.CreateTableDef(strTableName)
    
    'Establece las propiedades de la cadena Connect a la colección del nuevo objeto TableDef.
        tdfFront.Connect = ";DATABASE=" & strBack
        tdfFront.SourceTableName = strTableName
        
        dbsFront.TableDefs.Append tdfFront
            
        If Not tdfFront Is Nothing Then Set tdfFront = Nothing
    
        If Not dbsFront Is Nothing Then
            dbsFront.Close
            Set dbsFront = Nothing
        End If
       
End Sub

