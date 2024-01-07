'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-selectores-de-fechas/
'                     Destello formativo 399
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Aplicación tipo App
' Autor original    : Luis Viadel
' Creado            : 27/12/2023
' Propósito         : Hace unos días me envió un mensaje McPegasus diciéndome que no habíamos hecho un destello mostrando selectores de fecha 
'					  con trimestres. Así que hemos preparado este destello.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Código de la creación de los combos. Añadir en un formulario o crear una función
'Combo de años
    num = 0
    Set comb = Me.cboAno
        num = comb.ListCount
'Vaciamos el combo
        For i = 1 To num
            comb.RemoveItem 0
        Next i
        strSQL = "SELECT First(Year([Mi_fecha])) AS anofac, Count(Mis_Facturas.[Mi_fecha]) AS Repeticion " & _
                 "FROM Mis_Facturas " & _
                 "GROUP BY Year([Mi_fecha]) " & _
                 "HAVING (((Count(Mis_Facturas.[Mi_fecha]))>=1))"
        Set rstTable = CurrentDb.OpenRecordset(strSQL)
            Do Until rstTable.EOF
'Rellenamos de nuevo el combo
                comb.AddItem rstTable!anofac
            rstTable.MoveNext
            Loop
        Set rstTable = Nothing
        Set comb = Nothing
'Combo de trimestres
    num = 0
    Set comb = Me.cboTrim
        num = comb.ListCount
        For i = 1 To num
            comb.RemoveItem 0
        Next i
        strSQL = "SELECT First(Format([Mi_fecha],""" & "q" & """)) AS trimestre, Count(Mis_Facturas.[Mi_fecha]) AS Repeticion " & _
                "FROM Mis_Facturas " & _
                "GROUP BY Format([Mi_fecha],""" & "q" & """) " & _
                "HAVING (((Count(Mis_Facturas.[Mi_fecha]))>=1))"
        Set rstTable = CurrentDb.OpenRecordset(strSQL)
            Do Until rstTable.EOF
                comb.AddItem "T" & rstTable!trimestre
            rstTable.MoveNext
            Loop
        Set rstTable = Nothing
    Set comb = Nothing
'Combo de meses
    Set comb = Me.cboMes
        num = comb.ListCount
        For i = 1 To num
            comb.RemoveItem 0
        Next i
        strSQL = "SELECT First(monthname(month([Mi_fecha]))) AS mesfac, Count(Mis_Facturas.[Mi_fecha]) AS Repeticion " & _
                 "FROM Mis_Facturas " & _
                 "GROUP BY month([Mi_fecha]) " & _
                 "HAVING (((Count(Mis_Facturas.[Mi_fecha]))>=1))"
        Set rstTable = CurrentDb.OpenRecordset(strSQL)
            Do Until rstTable.EOF
                strItem = rstTable!mesfac
                comb.AddItem strItem
            rstTable.MoveNext
            Loop
        Set rstTable = Nothing
    Set comb = Nothing