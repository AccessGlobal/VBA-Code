Sub ControlarTabla()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-y-manipular-una-tabla-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ControlarTabla
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : septiembre 22
' Propósito         : trabajar con tablas temporales que se crean y eliminan en la misma rutina
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/database-execute-method-dao
'                   : https://learn.microsoft.com/es-es/office/vba/api/access.docmd.deleteobject?redirectedfrom=MSDN
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, seleccionar una fuente que no se encuentre
'                     en el sistema y pulsar F5 para ver su funcionamiento.
'
' Sub ControlarTabla_test()
'
'       Call Controlartabla
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dbs As Database

'Crea una tabla temporal donde guarda la configuración actual
    Set dbs = CurrentDb
        dbs.Execute "CREATE TABLE produold (idprodu INTEGER, producod CHAR);"

'Rellena la nueva tabla con los datos de otra tabla
        dbs.Execute " INSERT INTO produold SELECT idprodu, producod FROM [produ];"
        
'Vacíamos la tabla de nuevo
        dbs.Execute " DELETE * FROM produold;"
        dbs.Close
    Set dbs = Nothing
        
'Borra la tabla una vez que la hemos utilizado
        DoCmd.DeleteObject acTable, "produold"
    
End Sub