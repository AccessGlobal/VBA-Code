Public Sub ConfiguraUsuario()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-createtextfile
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConfiguraUsuario
' Autor             : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : marzo 21
' Propósito         : escribir datos de configuración de usuario para utilizarlos en algún momento
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/createtextfile-method
'                     https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/writeline-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso, archivo
Dim ruta As String

    ruta = "C:\...\configusuar.txt"

'Comprueba si ya existe el fichero
'Si el fichero existe en la carpeta, lo elimina para grabarlo con los datos actuales
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(ruta) Then
            Kill (ruta)
        End If
'Graba de nuevo el fichero
        Set archivo = fso.CreateTextFile(ruta, False)
            Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM tabla1 ORDER BY id")
                Do Until rstTable.EOF
                    archivo.WriteLine rstTable!Id & ", " & rstTable!Propiedad
                rstTable.MoveNext
                Loop
            Set rstTable = Nothing
            archivo.Close
        Set archivo = Nothing
    Set fso = Nothing

End Sub