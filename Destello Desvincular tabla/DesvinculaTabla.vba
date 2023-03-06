Sub mcDesvincular_mi_primera_tabla()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-desvincular-mi-primera-tabla-de-access/
'                     Destello formativo 279
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcDesvincular_mi_primera_tabla
' Autor original    : Rafael Andrada | McPegasus |https://beesoftware.es
' Colaborador       : Agradecimiento a Juan Luna por su contribución en la obtención de método refreshdatabasewindow
' Creado            : 07/01/2022
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : desvincular una tabla mediante VBA
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-object-dao
'                     https://learn.microsoft.com/es-es/office/vba/api/access.application.refreshdatabasewindow
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dbsFront  As DAO.Database
Dim strTableName As String

    Set dbsFront = CurrentDb
    
        strTableName = "Test"
            
        dbsFront.TableDefs.Delete (strTableName)
    
        Call Application.RefreshDatabaseWindow
    
    If Not dbsFront Is Nothing Then
        dbsFront.Close
        Set dbsFront = Nothing
    End If

End Sub