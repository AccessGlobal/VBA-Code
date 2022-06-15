Public Function ConstruirRuta(strRutafichero As String, strArchivo As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-BuildPath
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConstruirRuta
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Combina una ruta de carpeta y el nombre de una carpeta o archivo y devuelve la combinación con separadores de ruta válidos.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/buildpath-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckConstruirRuta_test()
' Dim ruta As String
' Dim nombrearchivo As String
 
'    ruta = "C:\MiPrograma"
'    nombrearchivo = "MiApp.txt"
'
'    Debug.Print ConstruirRuta(ruta, nombrearchivo)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
        ConstruirRuta = fso.BuildPath(strRutafichero, strArchivo)
        
    Set fso = Nothing
    
End Function
