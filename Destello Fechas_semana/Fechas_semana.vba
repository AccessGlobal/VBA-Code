Public Function mcdtmFindDateFirstOrLastDayOfTheWeek(ByVal dtmDateSearched As Date, ByVal bytQuarterInterval As Integer, ByVal blnFirstDay As Boolean, Optional ByVal blnDifferentiateMonth As Boolean) As Date
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/obtener-la-fecha-del-primer-dia-de-la-semana-o-ultimo-dia-de-la-semana-en-access-vba/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Obtener la fecha del primer día de la semana o último día de la semana en Access VBA
' Autor             : Rafael Andrada .:McPegasus:. | BeeSoftware
' Actualizado       : 24/03/2022
' Propósito         : Obtener la fecha que corresponda al inicio de la semana (lunes) según una fecha dada. También se puede obtener la fecha del último día de la semana (domingo)
' Retorno           : Un valor de tipo Date que nos indica la fecha de inicio de semana o último día según los parámetros indicados.
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     dtmDateSearched           Obligatorio             El valor Date especifica la fecha desde la que se parte para conseguir el día solicitado según el parámetro blnFirstDay.
'                     bytQuarterInterval        Obligatorio             El valor Integer especifica el intervalo de tiempo en meses que se desea obtener. 0 es para el mismo mes.
'                     blnFirstDay               Obligatorio             El valor Boolean especifica si se desea encontrar el primer o último día del periodo según el parámetro bytQuarterInterval.
'                     blnDifferentiateMonth     Obligatorio/Opcional    El valor String especifica el SUSTITUIR_POR_UNA_BREVE_DESCRIPCIÓN_DEL_PARÁMETRO
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Importante        : La primera fecha que se obtiene es la del mismo año, no del año anterior ni tampoco del año postarior. Si el día pasado es 03/01/2025, se obtiene como primer día de la semana 01/01/2025 y como último 05/01/2025.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  todo el procedimiento desde el Sub hasta el End Sub
'                     al portapapeles y pega en el editor de VBA de tu aplicación MS Access. Descomentar todas las líneas que nos interese (se aconseja seleccionar
'                     todas las líneas del ejemplo y utilizar el botón 'Bloque sin comentarios' de la barra de herramientas 'Edición').
'                     Pulsar F5 para ver su funcionamiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub mcdtmFindDateFirstOrLastDayOfTheWeek_test()
'   En el procedimiento adjunto se puede obtener diversos ejemplos de funcionamiento.
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    Dim intWork                                     As Integer
    Dim intAño                                      As Integer
    Dim intAñoOtro                                  As Integer
    Dim intMes                                      As Integer
    Dim intMesOtro                                  As Integer
    Dim dtmDíaSemanaOtroAño                         As Date
    'Conocer el día de la semana de la fecha pasada por parámetro.
    intWork = Weekday(dtmDateSearched, vbMonday)
    If blnFirstDay Then
        dtmDíaSemanaOtroAño = DateAdd("d", -intWork + 1, dtmDateSearched)
    Else
        dtmDíaSemanaOtroAño = DateAdd("d", 7 - intWork, dtmDateSearched)
    End If
    intAño = Year(dtmDateSearched)
    intAñoOtro = Year(dtmDíaSemanaOtroAño)
    intMes = Month(dtmDateSearched)
    intMesOtro = Month(dtmDíaSemanaOtroAño)
    If blnFirstDay Then
        'Obtener el primer día de la semana.
        If intAño = intAñoOtro Then
            If blnDifferentiateMonth Then
                If intWork = 1 Then
                    intWork = 0
                Else
                    'En caso de retornar el primer día de la ultima semana del propio mes.
                    If intMes = intMesOtro Then
                        intWork = (intWork - 1) * -1
                    Else
                        intWork = Weekday(DateSerial(intAño, intMes, 1), vbMonday) - intWork
                    End If
                End If
            Else
                If intWork = 1 Then
                    intWork = 0
                Else
                    intWork = (intWork - 1) * -1
                End If
            End If
        Else
            'En caso de ser la primera semana del año, el primer día es posible que sea del año anterior. Hay que obtener el día de la semana del siguiente al 31.
            intWork = intWork - (Weekday(DateSerial(intAñoOtro, 12, 31), vbMonday) + 1)
            intWork = intWork * -1
        End If
    Else
        'Obtener el primer día de la semana.
        If intAño = intAñoOtro Then
            If blnDifferentiateMonth Then
                'En caso de retornar el último día de la ultima semana del propio mes.
                If intMes = intMesOtro Then
                    '24/03/2022.
                    'En caso de ser la última semana de un mes, retornar el último día del mes al que pertenece el día según criterio.
                    intWork = 7 - intWork
                Else
                    '24/03/2022.
                    'En caso de ser la última semana de un mes, retornar el último día del mes al que pertenece el día según criterio.
                    intWork = Weekday(DateSerial(intAño, intMes + 1, 1), vbMonday) - intWork - 1
                End If
            Else
                '24/03/2022.
                'En caso de ser la última semana de un mes, retornar el último día del mes al que pertenece el día según criterio.
                intWork = 7 - intWork
            End If
        Else
            '24/03/2022.
            'En caseo de ser la última semana del año, el último día es posible que sea del año siguiente. Hay que obtener el día de la semana que corresponde al 31.
            'Obtener el número de día del último día del año.
            intWork = (Weekday(DateSerial(intAño, 12, 31), vbMonday)) - intWork
        End If
    End If
    mcdtmFindDateFirstOrLastDayOfTheWeek = DateAdd("d", intWork, dtmDateSearched)
End Function

Sub mcdtmFindDateFirstOrLastDayOfTheWeek_test()
    'Fechas de la primera semana del años en la que el primer día puede estar en un año diferente. En este caso el primer día de la semana es el 1.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#1/1/2024#, 0, True, True)         '01/01/2024         'El primer día es lunes.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#1/1/2024#, 0, True, False)        '01/01/2024         'El primer día es lunes.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#1/1/2024#, 0, False, True)        '07/01/2024
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#1/1/2024#, 0, False, False)       '07/01/2024
    'Fechas de la última semana del año donde el último día puede estar en un año diferente. En este caso el último día de la semana es 31.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#12/31/2024#, 0, True, True)       '30/12/2024         'El último día es martes.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#12/30/2024#, 0, False, True)      '31/12/2024
    'Fechas con días de febrero para comprobar los días 28 y 29.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#2/28/2023#, 0, True, True)        '27/02/2023         'El último día es martes.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#2/27/2023#, 0, False, True)       '28/02/2023
    '
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#2/28/2024#, 0, True, True)        '26/02/2023         'El último día es martes.
    '    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#2/27/2024#, 0, False, True)       '29/02/2023
    'Otras fechas cualesquiera.
    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#8/15/2024#, 0, True, True)        '12/08/2024
    Debug.Print mcdtmFindDateFirstOrLastDayOfTheWeek(#8/16/2024#, 0, False, True)       '18/08/2024
End Sub
