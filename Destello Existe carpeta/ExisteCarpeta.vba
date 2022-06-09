Public Function ExisteCarpeta(ByVal ruta As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-folderexists
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExisteCarpeta
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : saber si una determinada carpeta existe
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/folderexists-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
' Dim ruta as string
'
' ruta = "C:\...\MiCarpeta\"
'
'        Debug.print ExisteCarpeta(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")
    
        ExisteCarpeta = fso.FolderExists(ruta)
        
     Set fso = Nothing

End Function
