Public Function ObtenerExtension(ByVal strRutafichero As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-GetExtensionName
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ObtenerExtension
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Obtener la extension de un fichero, tanto si pasamos la ruta completa, como si solo pasamos el nombre del fichero.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getextensionname-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
' Dim ruta as string
'
' ruta = "C:\...\MiCarpeta\MiArchivo.txt"
'
'        Debug.print ObtenerExtension(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
        ObtenerExtension = fso.GetExtensionName(strRutafichero)
        
    Set fso = Nothing
    
End Function
