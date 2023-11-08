Sub ConsultaCompleja()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-consultas-complejas
'                   : Destello formativo 374
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConsultaDeLotes
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : Noviembre 22
' Propósito         : trabajar con tablas temporales para crear consultas complejas
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://access-global.net/vba-crear-y-manipular-una-tabla-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, seleccionar una fuente que no se encuentre
'                     en el sistema y pulsar F5 para ver su funcionamiento.
'
' Sub ConsultaCompleja_test()
'
'       Call ConsultaCompleja
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dbs As DAO.Database
Dim rstTable As DAO.Recordset
Dim idarticulo As Integer


'Borra la tabla temporal si existe. Si no existe lanza un error que libramos con el control de errores
    On Error Resume Next
    
    DoCmd.DeleteObject acTable, "ArticulosRequeridosTmp"
    
    On Error GoTo 0
   
'Crea laa tabla temporal
    Set dbs = CurrentDb
        dbs.Execute "CREATE TABLE ArticulosRequeridosTmp (idarticulo INTEGER, Lote VARCHAR(4),Estado YESNO);"
'Rellena la nueva tabla con los datos de consulta
'Artículos lotes abiertos
                dbs.Execute "INSERT INTO ArticulosRequeridosTmp (idarticulo, Lote, Estado) SELECT Articulo_5.idarticulo,Articulos_Lotes_5.Lote, Articulos_Lotes_5.Estado " & _
                                        "FROM Articulo_5 INNER JOIN Articulos_Lotes_5 ON Articulo_5.idArticulo = Articulos_Lotes_5.idArticulo " & _
                                        "WHERE (((Articulos_Lotes_5.Estado)=True));"
'Artículos lotes cerrados
                dbs.Execute "INSERT INTO ArticulosRequeridosTmp (idarticulo) SELECT Articulos_Lotes_5.idarticulo " & _
                            "FROM Articulos_Lotes_5 " & _
                            "GROUP BY Articulos_Lotes_5.idarticulo, Articulos_Lotes_5.Estado " & _
                            "HAVING (((Articulos_Lotes_5.Estado)=False));"

'Artículos sin lote
                dbs.Execute "INSERT INTO ArticulosRequeridosTmp (idarticulo, Estado) SELECT Articulo_5.idArticulo, Articulos_Lotes_5.Estado " & _
                            "FROM Articulo_5 LEFT JOIN Articulos_Lotes_5 ON Articulo_5.idArticulo = Articulos_Lotes_5.idarticulo " & _
                            "WHERE (((Articulos_Lotes_5.Estado) Is Null));"

'Recorremos el recordset para eliminar líneas en blanco generadas por los lotes terminados en artículos con lotes abiertos.
        Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM ArticulosRequeridosTmp ORDER BY idarticulo, Lote DESC")
            Do Until rstTable.EOF
                If rstTable!idarticulo = idarticulo And IsNull(rstTable!LOTE) And rstTable!estado = False Then
                    rstTable.Delete
                    GoTo LinNext
                End If
                idarticulo = rstTable!idarticulo
LinNext:
            rstTable.MoveNext
            Loop
        Set rstTable = Nothing
        dbs.Close
    Set dbs = Nothing
        
End Sub