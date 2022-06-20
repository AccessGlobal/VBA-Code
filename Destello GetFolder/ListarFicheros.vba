Public Function ListarFicheros(strCarpeta) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-getfolder
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListarFicheros
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Recorre la carpeta y devuelve la lista de ficheros.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub BorraCopiaTemp_test()
' Dim ruta As String
'
'    ruta = "C:\MiPrograma\"
'
'    Debug.Print GetFileName(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim objCarpeta As Object
Dim objFile As Object
Dim strSalida As String
Dim colFiles

    Set fso = CreateObject("Scripting.FileSystemObject")
                   
        Set objCarpeta = fso.GetFolder(strCarpeta)
        
            Set colFiles = objCarpeta.Files
            
                For Each objFile In colFiles
                    strSalida = strSalida & objFile.Name & vbCrLf
                Next
                
                ListarFicheros = strSalida
                
            Set colFiles = Nothing
        
        Set objCarpeta = Nothing
    
    Set fso = Nothing
    
End Function