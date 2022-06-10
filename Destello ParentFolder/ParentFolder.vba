Public Function ObtenerCarpeta(ByVal strRutafichero As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-GetParentFolderName
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ObtenerCarpeta
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Obtener el nombre de la carpeta donde está un fichero, pasando la ruta completa, nos devuelve la ruta completa
'                     Mediante la función Split, obtenemos el nombre de la carpeta que lo contiene.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/GetParentFolderName-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
' Dim ruta as string
'
' ruta = "C:\...\MiCarpeta\MiArchivo.vba"
'
'        Debug.print ObtenerCarpeta(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim nmat As Variant
Dim n As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    
        nmat = Split(fso.GetParentFolderName(strRutafichero), "\")
        
        For n = 0 To UBound(nmat)
            ObtenerCarpeta = nmat(n)
        Next n
            
    Set fso = Nothing
    
End Function
