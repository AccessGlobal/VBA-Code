Public Function ConvertirHoraEnDecimal(ByVal Hora As Date) As Single
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-convertir-hora-en-decimal
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConvertirHoraEnDecimal
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : enero 2013
' Propósito         : convertir un formato hora en un formato decimal
' Retorno           : nos devuelve un número (single) con el valor de la hora
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Hora         Obligatorio      hora que queremos convertir
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub ConvertirHoraEnDecimal_test()
'
'        Debug.Print ConvertirHoraEnDecimal("11:30")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim TB As Variant

TB = Split(Hora, ":")

ConvertirHoraEnDecimal = TB(0) + ((TB(1) * 100) / 60) / 100

End Function