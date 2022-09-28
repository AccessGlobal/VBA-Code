Option Compare Database
Option Explicit

Sub ControlarTablaSQL()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-y-manipular-una-tabla-en-tiempo-de-ejecucion-runsql
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ControlarTablaSQL
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : septiembre 22
' Propósito         : trabajar con tablas temporales que se crean y eliminan en la misma rutina
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/es-es/office/vba/api/access.docmd.runsql
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, seleccionar una fuente que no se encuentre
'                     en el sistema y pulsar F5 para ver su funcionamiento.
'
' Sub ControlarTabla_test()
'
'       Call ControlartablaSQL
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Crea una tabla temporal donde guarda la configuración actual
        DoCmd.RunSQL "CREATE TABLE produold (idprodu INTEGER, producod CHAR)"

'Rellena la nueva tabla con los datos de otra tabla
        DoCmd.RunSQL " INSERT INTO produold SELECT idprodu, producod FROM [produ]"
        
'Vacíamos la tabla de nuevo
        DoCmd.RunSQL " DELETE * FROM produold"
        
'Borra la tabla una vez que la hemos utilizado
        DoCmd.DeleteObject acTable, "produold"
    
End Sub
