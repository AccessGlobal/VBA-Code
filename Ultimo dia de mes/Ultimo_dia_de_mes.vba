Public Function mcdtmFindDateLastDayofMonth(ByVal dtmDateSearched As Date, ByVal intMonthsInterval As Integer) As Date

    Dim dtmEndOfMonth                           As Date             'Para conocer el mes siguiente según los parámetros introducidos.
    Dim dtmNextMonth                            As Date             'Para conocer el último día del mes según los parámetros introducidos.
    
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/conocer-el-ultimo-dia-de-un-mes-mcdtmfinddatelastdayofmonth/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcdtmFindDateLastDayofMonth.
' Autor original    : http://support.microsoft.com/kb/103184/es (FindEOM Function: This function takes a date as an argument and returns the last day of the month.
' Adaptado por      : Rafael Andrada .:McPegasus:. | BeeSoftware.
' Actualizado       : 14/03/2014.
' Propósito         : Conocer el último día de un mes según los argumentos pasados.
' Retorno           : Una fecha que corresponde a la última del mes según los argumentos pasados.
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     dtmDateSearched       Obligatorio     El valor Date especifica la fecha desde la que se parte para conseguir el último día del mes en caso de bytMesesIntervalo = 1.
'                     intMonthsInterval     Obligatorio     El valor Byte especifica el contiene el intervalo de tiempo en meses que se desea agregar o disminuir.
'                     intMonthsInterval     Opciones        En caso de 0 retorna el último día del mes anteior. Si se desea el último día del mes actual el valor es 1.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar todo el procedimiento desde el Sub hasta el End Sub
'                     al portapapeles y pega en el editor de VBA de tu aplicación MS Access. Descomentar todas las líneas que nos interese (se aconseja seleccionar
'                     todas las líneas del ejemplo y utilizar el botón 'Bloque sin comentarios' de la barra de herramientas 'Edición').
'                     Pulsar F5 para ver su funcionamiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub mcdtmFindDateLastDayofMonth_test()
'    Debug.Print mcdtmFindDateLastDayofMonth(Date, 2)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------      
    
    dtmNextMonth = DateAdd("m", intMonthsInterval + 1, dtmDateSearched)     '27/12/2022 = ("m", 11, 27/01/2020)
    dtmEndOfMonth = dtmNextMonth - DatePart("d", dtmNextMonth)              '30/11/2022 = 27/12/2022 - ("d", 27/12/2022) -> 31/03/2020 = 27/12/2022 - 27
    mcdtmFindDateLastDayofMonth = dtmEndOfMonth

End Function