Public Function AS_EsBisiesto(optional IntAño As integer = 0) As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/saber-si-un-ano-es-bisiesto/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AS_EsBisiesto
' Autor original    : Microsoft
' Adaptado por      : Alba Salvá | Isis
' Actualizado       : 19/08/2001
' Propósito         : Saber si un año es bisiesto o no según el argumento pasado.
' Retorno           : Verdadero o Falso, según el argumento pasado.
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     IntAño       Opcional    El valor entero del año a verificar
'                     
'-----------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub AS_EsBisiesto_test()
'    Debug.Print AS_EsBisiesto(2000)' Debe devolver Verdadero
'    Debug.Print AS_EsBisiesto(2020)' Debe devolver Verdadero
'    Debug.Print AS_EsBisiesto(2022)' Debe devolver Falso
'    Debug.Print AS_EsBisiesto' Debe devolver Verdadero	o Falso dependiendo del año actual'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------- 

	If IntAño = 0 Then IntAño = Year(Date)
 	If Day(DateSerial(IntAño, 3, 0)) = 29 Then AS_EsBisiesto = True
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------