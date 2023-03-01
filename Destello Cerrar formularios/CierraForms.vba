Sub CierraForms()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-cerrar-todos-los-formularios
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CierraForms
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : cerrar todos los formularios que estén abiertos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/api/access.allforms
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim obj As Object
    For Each obj In Application.CurrentProject.AllForms
        If obj.IsLoaded = True Then
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
End Sub