Public Function TieneControlErrores(vbc As VBIDE.VBComponent, lngInicioProc As Long) As Integer
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/mis-procedimientos-tienen-tratamiento-de-errores
'--------------------------------------------------------------------------------------------------------
' Título            : TieneControlErrores
' Autor original    : Alba Salvá
' Creado            : 2023
' Propósito         : detectar los procedimientos que no contienen tratamiento de errores
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo           Descripción
'-------------------------------------------------------------------------------------------------------
'                     vbc               Obligatorio    VBIDE
'                     lngInicioProc     Obligatorio    Nombre de la tabla
'--------------------------------------------------------------------------------------------------------
' Retorno           : Long
'                     0 = No tiene tratamiento de errores
'                     1 = tratamiento comentado
'                     2 = tratamiento disponible
'--------------------------------------------------------------------------------------------------------
Dim lngFinLineaProc As Long
Dim lngInicioLineaTr As Long
Dim lngFinLineaTr As Long
Dim lngFinColTr As Long
Dim lngProcTyp As Long
Dim PosManejaError As Boolean

    On Error GoTo lbError
    
    With vbc.CodeModule
        lngFinLineaProc = lngInicioProc + .ProcCountLines(.ProcOfLine(lngInicioProc, _
        lngProcTyp), lngProcTyp) - 1 ' Calculo la última línea del procedimiento.
        lngInicioLineaTr = lngInicioProc
        PosManejaError = .Find("On Error", lngInicioLineaTr, 0, lngFinLineaTr, lngFinColTr, True) ' Busco la aparición de "On Error"
        Do While PosManejaError = True And (lngFinLineaTr < lngFinLineaProc)
            If Not EsComentario(.Lines(lngFinLineaTr, 1)) Then
                TieneControlErrores = 2
                Exit Function
            Else
                TieneControlErrores = 1
                Exit Function
            End If
            lngInicioLineaTr = lngInicioLineaTr + 1 ' Si no se ha encontrado, incremento el número de línea y sigo buscando
            PosManejaError = .Find("On Error", lngInicioLineaTr, 0, lngFinLineaTr, lngFinColTr, True)
        Loop
    End With

    GoTo lbFinally

lbError:

    MsgBox "Ha ocurrido un error" & vbCrLf & vbCrLf & _
           "Código de Error : " & Err.Number & vbCrLf & _
           "Origen del Error : ObtenerCode" & vbCrLf & _
           "Descripción Error : " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Línea No: " & Erl) _
           , vbOKOnly + vbCritical, "¡Ha ocurrido un error"

lbFinally:

End Function

Public Function EsComentario(strLinea As String) As Boolean
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/mis-procedimientos-tienen-tratamiento-de-errores
'--------------------------------------------------------------------------------------------------------
' Título            : EsComentario
' Autor original    : Alba Salvá
' Creado            : 2023
' Propósito         : comprobar si una línea de un procedimiento es un comentario
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo           Descripción
'-------------------------------------------------------------------------------------------------------
'                     strLinea          Obligatorio    Línea que se va a analizar
'--------------------------------------------------------------------------------------------------------
' Retorno           : booleano
'                     sí = es un comentario
'                     no = no es un comentario
'--------------------------------------------------------------------------------------------------------
    strLinea = Trim$(strLinea)

    If Left$(strLinea, 1) = "'" Or Left$(strLinea, 3) = "Rem" Then
        EsComentario = True
    End If

End Function

