Option Compare Database
Option Explicit

Public Function SemanaFecha(ByVal fechatest As Date) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-en-que-semana-estamos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : SemanaFecha
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : septiembre 2021
' Propósito         : conocer el número de semana al que pertenece una fecha cualquiera
' Retorno           : número de semana a la que pertenece la fecha
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     fechasemana     Obligatorio      fecha sobre la que queremos calcular el número
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datepart-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
'
'        Debug.Print SemanaFecha(date) 'Nos indica en la semana que nos encontramos hoy
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

SemanaFecha = DatePart("ww", fechatest, vbMonday, vbFirstFourDays)

End Function
