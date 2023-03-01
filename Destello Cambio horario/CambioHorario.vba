Option Compare Database
Option Explicit

Public Function CambioHorarioInvierno(ByVal datDia As Date) As Date
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-horario-de-verano
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CambioHorarioInvierno
' Autor original    : Emilio Sancha
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : Conocer la fecha de cambio de horario de invierno(ultimo domingo de Octubre)
' Retorno           : Valor date con el día que cambia la hora
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     datDia      Obligatorio      Número de días del mes (31)
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub CambioHorarioInvierno_test()
'
'Debug.Print CambioHorarioInvierno(#1/1/2022#)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim bytDia As Byte
Dim intAño As Integer

intAño = Year(datDia)

bytDia = 31
datDia = DateSerial(intAño, 10, bytDia)

' busco el ultimo domingo (primero empezando por el final)
Do While Weekday(datDia) <> vbSunday
    datDia = DateSerial(intAño, 10, bytDia)
    bytDia = bytDia - 1
Loop

CambioHorarioInvierno = datDia

End Function


Public Function CambioHorarioVerano(ByVal datDia As Date) As Date
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-horario-de-verano
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CambioHorarioVerano
' Autor original    : Emilio Sancha
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : Conocer la fecha de cambio de horario de verano(ultimo domingo de marzo)
' Retorno           : Valor date con el día que cambia la hora
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     datDia      Obligatorio      Número de días del mes (31)
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub CambioHorarioVerano_test()
'
'Debug.Print CambioHorarioVerano(#1/1/2022#)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim bytDia As Byte
Dim intAño As Integer

intAño = Year(datDia)

bytDia = 31
datDia = DateSerial(intAño, 3, bytDia)

' busco el ultimo domingo (primero empezando por el final)
Do While Weekday(datDia) <> vbSunday
    datDia = DateSerial(intAño, 3, bytDia)
    bytDia = bytDia - 1
Loop

CambioHorarioVerano = datDia

End Function


Public Function CalculoGMT(ByVal datDia As Date) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CalculoGMT
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Propósito         : Conocer el GMT en el que nos encontramos
' Retorno           : Valor string con la descripción del GMT actual
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     datDia      Obligatorio      Fecha en la que queremos saber su GMT
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub CalculoGMT_test()
'Dim gmt as date
'
'gmt= CalculoGMT(MiFecha)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

If datDia > CambioHorarioInvierno(datDia) And datDia < CambioHorarioVerano(datDia) Then
    GMT = "+01:00"
Else
    GMT = "+02:00"
End If
            
End Function

Sub CambioHorarioVerano_test()

Debug.Print CambioHorarioVerano(#1/1/2022#)

End Sub
