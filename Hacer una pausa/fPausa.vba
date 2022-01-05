Public Function fPausa(Optional ByRef seg As Long)

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    ' Fuente            : https://access-global.net/hagamos-una-pausa/
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    ' Título            : Pausa
    ' Autor original    : Desconocido
    ' Adaptado por      : Luis Viadel
    ' Actualizado       : marzo 2018
    ' Propósito         : Conseguir que el código se detenga durante un tiempo determinado
    ' Retorno           : No hay retorno
    ' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
    '                     Parte                 Modo           Descripción
    '---------------------------------------------------------------------------------------------------------------------- -------------------------
    '                     seg                  Opcional       El valor long especifica un número de segundos
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim SegInicio As Long, SegFin As Long
    
    On Error GoTo LinErr
    
    SegInicio = Timer
    SegFin = Timer + seg
        
    Do While Timer < SegFin
        DoEvents 'Corrección por si hay un cambio de día (86400=segundos/día)
            If Timer < SegInicio Then SegFin = SegFin - 86400
    Loop
    
    'Si se produce un error sale
    LinErr:
    
    End Function