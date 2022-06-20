Public Function ListarSubCarpetas(strCarpeta) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-getfolder-subfolders
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListarSubCarpetas
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Recorre la carpeta y devuelve la lista de subcarpetas.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub ListarSubCarpetas_test()
' Dim ruta As String
'
'    ruta = "C:\MiPrograma\subcarpeta1\subcarpeta2"
'
'    Debug.Print ListarSubCarpetas(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim objCarpeta As Object
Dim objSubCarpeta As Object
Dim strSalida As String
Dim colSubCarpetas

    Set fso = CreateObject("Scripting.FileSystemObject")
                   
        Set objCarpeta = fso.GetFolder(strCarpeta)
        
            Set colSubCarpetas = objCarpeta.SubFolders
        
                For Each objSubCarpeta In colSubCarpetas
        
                    strSalida = strSalida & objSubCarpeta.Name & vbCrLf
        
                Next
                
                ListarSubCarpetas = strSalida
                
            Set colSubCarpetas = Nothing
        
        Set objCarpeta = Nothing
    
    Set fso = Nothing
    
End Function