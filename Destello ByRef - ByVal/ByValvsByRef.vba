'Copia y pega este código en un módulo estándar
Option Compare Database
Option Explicit

Sub ByValvsByRef()
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-byval-vs-byref
'                     Destello formativo 273
' Fuente original   : https://learn.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ByValvsByRef
' Autor original    : Microsoft
' Creado            : 15/09/2021
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : entender la diferencia entre ByVal y ByRef. Utiliza el procedimiento calcular para realizar los cálculos necesarios
' Argumentos        : La sintaxis de la función"calcular" consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                     tasa             Obligatorio    Valor que pasamos como ByVal
'                     deuda            Obligatorio    Valor que pasamos como ByRef
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Cómo funciona     : dependiendo de cómo pasemos cada uno de los argumentos de la función, podemos obtener resultados diferentes
' Referencias       : https://learn.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/procedures/differences-between-passing-an-argument-by-value-and-by-reference
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Test              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el siguiente bloque a portapapeles y pégalo
'                     en el editor de VBA. Descomenta la línea de te interese y pulsa F5 para ver su funcionamiento.
' Para realizar el test, camba la forma de pasar el argumento en la función "calcular"
' Test 1
'Sub calcular(ByVal tasa As double, ByRef deuda As Double)
'    deuda = deuda + (deuda * tasa / 100)
'End Sub
'
' Test 2
'Sub calcular(ByVal tasa As doubLE, ByVal deuda As Double)
'    deuda = deuda + (deuda * tasa / 100)
'End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim TasaInferior As Double
Dim ValorInicial As Double
Dim deudaConIntereses As Double
Dim deudaStr As String
Const TasaSuperior As Double = 12.5

    ValorInicial = 4999.99
    
    TasaInferior = TasaSuperior * 0.6
    
    deudaConIntereses = ValorInicial * TasaSuperior
    
'Calcula la deuda total con la tasa de interés alta aplicada.
'Argumento TasaSuperior es una constante, lo cual es apropiado para un parámetro ByVal.
'El argumento DeudaConIntereses debe ser una variable porque el procedimiento cambiará su valor al calculado
    deudaStr = Format(deudaConIntereses, "currency")
    
    Debug.Print "Deuda con interés alto: " & deudaStr

'Repite el proceso con la TasaInferior. El argumento TasaInferior no es una constante, pero el parámetro ByVal lo protege  de accidentes
'o cambios intencionales en el procedimiento.

'Volvemos a establecer la deuda con intereses en el valor original
    deudaConIntereses = ValorInicial * TasaInferior
    
    Call calcular(TasaInferior, deudaConIntereses)
    
'Cambiamos a tipo de moneda
    deudaStr = Format(deudaConIntereses, "currency")
    
    Debug.Print "Deuda con interés bajo: " & deudaStr
    
End Sub

'La tasa de parámetro es un parámetro de ByVal porque el procedimiento no cambia el valor del argumento correspondiente en el
'código de llamada

'Si embargo, el valor calculado del parámetro de la deuda debe ser reflejado en el calor del argumento correspondiente en el
'código de llamada. Por lo tanto debe declararse ByRef.

Sub calcular(ByVal tasa As Double, ByRef deuda As Double)

    deuda = deuda + (deuda * tasa / 100)
    
End Sub
