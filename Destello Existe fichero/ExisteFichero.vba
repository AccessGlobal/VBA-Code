Public Function ExisteFichero(ByVal ruta As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-fileexists
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExisteFichero
' Autor             : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : junio 21
' Propósito         : saber si un determinado fichero existe
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fileexists-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
' ruta = "C:\...\MiFichero.txt"
'
'        Debug.print ExisteFichero(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso, Archivo

'Comprueba si ya existe el fichero
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(ruta) Then
            ExisteFichero = True
        Else
            ExisteFichero = False
        End If
    Set fso = Nothing

End Function