Public Sub mcOcultar_mi_primera_tabla_PERO_OCULTAR_DE_VERDAD()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-donde-esta-mi-tabla
'                     Destello formativo 280
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcOcultar_mi_primera_tabla_PERO_OCULTAR_DE_VERDAD
' Autor original    : Rafael Andrada | McPegasus |https://beesoftware.es
' Colaborador       : A mi estimable compañero Abelardo Ramírez por compartir lo que sabe y lo que le pregunto.
' Creado            : febrero 2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : ocultar una tabla mediante VBA
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-object-dao
'                     https://learn.microsoft.com/en-us/cpp/mfc/reference/cdaotabledefinfo-structure?view=msvc-170
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dbsFront As DAO.Database
Dim strTableName As String
    
    Set dbsFront = CurrentDb
    
        strTableName = "Test"
        
        dbsFront.TableDefs("Test").Attributes = dbHiddenObject  '1
        
        Call Application.RefreshDatabaseWindow

    If Not dbsFront Is Nothing Then
        dbsFront.Close
        Set dbsFront = Nothing

    End If

End Sub
