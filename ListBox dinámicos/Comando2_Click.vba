'----------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/listbox-dinamicos/
'----------------------------------------------------------------------------------------
' Título            : listbox dinámicos
' Autor original    : Antonio Otero
' Propósito         : Rellenar listbox con los campos de una trabla, de forma dinámica
'Coloca este código en el código de tu formulario
'----------------------------------------------------------------------------------------
Private Sub Comando2_Click()

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    
    Set dbs = CurrentDb
        Me.l_campos.RowSource = ""
        Set tdf = dbs.TableDefs("CLIENTES")
            For Each fld In tdf.Fields
                Me.l_campos.AddItem fld.Name & ";" & fld.Type
            Next
        Set fld = Nothing
        Set tdf = Nothing
    Set dbs = Nothing
    
End Sub