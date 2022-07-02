Public Function ListarFicherosPropiedades(strCarpeta) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-coleccion-files
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListarFicherosPropiedades
' Autor             : Luis Viadel
' Fecha             : junio 2022
' Propósito         : Recorre la carpeta y devuelve la lista de ficheros y sus propiedades
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'                     https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/files-collection
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub ListarFicherosPropiedades_test()
' Dim ruta As String
'
'    ruta = "C:\MiPrograma\"
'
'    Debug.Print ListarFicherosPropiedades(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object, objCarpeta As Object, objFile As Object
Dim strSalida As String
Dim colFiles
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        
        Set objCarpeta = fso.GetFolder(strCarpeta)
            
            Set colFiles = objCarpeta.Files
                For Each objFile In colFiles
                    strSalida = strSalida & "Nombre: " & objFile.Name & vbCrLf
                    strSalida = strSalida & "Atributos: " & objFile.Attributes & vbCrLf
                    strSalida = strSalida & "Fecha creación: " & objFile.DateCreated & vbCrLf
                    strSalida = strSalida & "Fecha último acceso: " & objFile.DateLastAccessed & vbCrLf
                    strSalida = strSalida & "Fecha última modificación: " & objFile.DateLastModified & vbCrLf
                    strSalida = strSalida & "Unidad: " & objFile.Drive & vbCrLf
                    strSalida = strSalida & "Carpeta principal: " & objFile.ParentFolder & vbCrLf
                    strSalida = strSalida & "Ruta: " & objFile.Path & vbCrLf
                    strSalida = strSalida & "Nombre corto: " & objFile.ShortName & vbCrLf
                    strSalida = strSalida & "Ruta corta: " & objFile.ShortPath & vbCrLf
                    strSalida = strSalida & "Tamaño: " & objFile.Size & " bytes" & vbCrLf
                    strSalida = strSalida & "Tipo: " & objFile.Type & vbCrLf
                    strSalida = strSalida & vbCrLf
                Next
                ListarFicherosPropiedades = strSalida
            
            Set colFiles = Nothing
        
        Set objCarpeta = Nothing
    
    Set fso = Nothing

End Function