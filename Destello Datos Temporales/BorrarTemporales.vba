Public Sub BorrarTemporales()
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-datos-temporales/
'                     Destello formativo 371
'--------------------------------------------------------------------------------------------------------
' Título            : BorrarTemporales
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : mayo 2010
' Propósito         : localizar las tablas temporales y vaciarlas de contenido
' Más información   : reglas para determinar s una tabla es temporal
'                     1. Son tablas de Access en local
'                     2. Utilizo el sufijo "temp" en el nombre
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA.
'
'Sub BorrarTemporales_Test()
'
'   Call BorrarTemporales
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim Tabla As String
Dim rstTable As DAO.Recordset
Dim i As Integer
Dim cPerson As New clsMsgBoxPersonalizado
Dim msgCabecera As String, msgTxt As String

'Recorro la tabla de sistema MSysObjects para obtener todas las tablas de la aplicación.
'Solo me interesan las que NO están conectadas mediante ODBC, son las de tipo 6. Identifico las temporales mediante el sufijo "temp"

    Set rstTable = CurrentDb.OpenRecordset("SELECT MSysObjects.Type, MSysObjects.ForeignName, MSysObjects.Name " & _
                                           " From MSysObjects " & _
                                           " WHERE (((MSysObjects.Type)=6) AND ((MSysObjects.ForeignName) Is Not Null));")
        Do Until rstTable.EOF
        
            Tabla = rstTable!ForeignName
            
'Cuando encuentro una tabla de Access, compruebo si el nombre contiene el sufijo "temp"
            If Tabla Like "*temp" Then
                i = i + 1 'Cuento las tablas que voy limpiando
'Borro el contenido de la tabla
                CurrentDb.Execute "DELETE * FROM " & Tabla, dbFailOnError

            End If
        
        rstTable.MoveNext
        Loop
    Set rstTable = Nothing

'Lanzo un mensaje mediante un mensaje personalizado como vimos en la clase magistral https://access-global.net/msgbox-personalizado/

    msgCabecera = "Proceso de borrado"
    msgTxt = "Se han vaciado " & i & " tablas temporales del sistema."

    With cPerson
        .Initialize 2, 3, msgCabecera, msgTxt
    End With
    

End Sub
