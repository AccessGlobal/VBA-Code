Public Function mcdtmFindDateFirstDayofMonth(ByVal dtmDateSearched As Date, ByVal intMonthsInterval As Integer) As Date
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/conocer-el-primer-dia-de-un-mes-mcdtmfinddatefirstdayofmonth/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcdtmFindDateFirstDayofMonth.
' Autor original    : Rafael Andrada .:McPegasus:. | BeeSoftware.
' Actualizado       : 27/01/2022.
' Propósito         : Conocer el primer día de un mes según los argumentos pasados.
' Retorno           : Una fecha que corresponde al primer día del mes según los argumentos pasados.
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     dtmDateSearched       Obligatorio     El valor Date especifica la fecha desde la que se parte para conseguir el primer día del mes en caso de bytMesesIntervalo = 1.
'                     intMonthsInterval     Obligatorio     El valor Byte especifica el contiene el intervalo de tiempo en meses que se desea agregar o disminuir.
'                     intMonthsInterval     Opciones        En caso de 0 retorna el primer día del mes actual.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  todo el procedimiento desde el Sub hasta el End Sub
'                     al portapapeles y pega en el editor de VBA de tu aplicación MS Access. Descomentar todas las líneas que nos interese (se aconseja seleccionar
'                     todas las líneas del ejemplo y utilizar el botón 'Bloque sin comentarios' de la barra de herramientas 'Edición').
'                     Pulsar F5 para ver su funcionamiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub mcdtmFindDateFirstDayofMonth_test()
'
'    Dim dtmDateSearched                             As Date
'
'
'    dtmDateSearched = #1/27/2021#
'    Debug.Print "Se parte de la fecha: " & dtmDateSearched & ". Primer día del mes según selección: " & mcdtmFindDateFirstDayofMonth(dtmDateSearched, 12) & ". Debe de salir: 01/01/2022"
'
'    dtmDateSearched = #1/27/2022#
'    Debug.Print "Se parte de la fecha: " & dtmDateSearched & ". Primer día del mes según selección: " & mcdtmFindDateFirstDayofMonth(dtmDateSearched, 0) & ". Debe de salir: 01/01/2022"
'
'    dtmDateSearched = #1/27/2023#
'    Debug.Print "Se parte de la fecha: " & dtmDateSearched & ". Primer día del mes según selección: " & mcdtmFindDateFirstDayofMonth(dtmDateSearched, -12) & ". Debe de salir: 01/01/2022"
'
'    Debug.Print
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dtmWork                                 As Date             'Para obtener un valor transitorio durante la ejecución del código.
dtmWork = DateSerial(Year(dtmDateSearched), Month(dtmDateSearched) + intMonthsInterval, 1)
mcdtmFindDateFirstDayofMonth = dtmWork
End Function