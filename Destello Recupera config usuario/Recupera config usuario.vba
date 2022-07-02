Public Sub RecuperaConfiguraUsuario()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-declaracion-open
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConfiguraUsuario
' Autor             : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : marzo 21
' Propósito         : recuperar los datos de configuración de usuario de un fichero de texto
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ruta As String, strLinea As String
Dim id As Integer
Dim IntWhere As Integer

ruta = "C:\Cow Technologies\Access global\Destellos formativos\Destello 179\configusuar.txt"

Open ruta For Input As #1
    While Not EOF(1)
        Line Input #1, strLinea
        IntWhere = InStr(strLinea, ",")
'Graba los datos del usuario en la tabla2
        Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM tabla2")
            rstTable.AddNew
                rstTable!id = left(strLinea, IntWhere - 1)
                rstTable!propiedad = right(strLinea, Len(strLinea) - IntWhere - 1)
            rstTable.Update
            rstTable.Close
        Set rstTable = Nothing
    Wend
Close #1

End Sub
