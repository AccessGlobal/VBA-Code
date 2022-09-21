Private Sub comboorigen_AfterUpdate()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-combos-dependientes
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : comboorigen_AfterUpdate
' Autor             : Luis Viadel
' Fecha             : sep 2022
' Propósito         : carga un segundo combo con la selección del primero (igual origen de datos)
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim comb As ComboBox
Dim num As Integer, i As Integer
Dim strItem As String

    num = 0
'Asignamos al combo a una variable que nos permita manejarlo
    Set comb = Me.combodestino
        num = comb.ListCount
'Borramos los datos que pudiese contener el combo de destino. Necesario en selecciones consecutivas
        For i = 1 To num
            comb.RemoveItem 0
        Next i
'Dimensionamos el combo con dos columnas. La primera contendrá el Id que nos permitirá hacer las restricciones
        With comb
            .ColumnCount = 2
            .ColumnWidths = "0;80"
        End With
        
        num = 0
'Recorremos la tabla origen de datos excepto el registro que ya hemos seleccionado en el combo 1
        Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM tabla3 WHERE idcombo<>" & Me.comboorigen)
            Do Until rstTable.EOF
                strItem = rstTable!idcombo & ";" & rstTable!ComboOpcion
'Vamos añadiendo los registros al combo
                comb.AddItem strItem, num
                num = num + 1
            rstTable.MoveNext
            Loop
        Set rstTable = Nothing
    Set comb = Nothing

End Sub