Sub InfoErrores()
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/tratamiento-de-errores-objeto-err
'--------------------------------------------------------------------------------------------------------
' Título            : InfoErrores
' Autor original    : Luis Viadel
' Creado            : marzo 2023
' Propósito         : conocer todos los atributos y métodos del objeto err
' Cómo funciona     : ante un error 94 - Uso no válido de null provocado por nosotros, mostramos
'                     todos sus atributos en la ventana de inmediato.
'                     Para ver el funcionamiento de clear: descomenta las 3 últimas líneas
'                     Para ver el funcionamiento de raise: comenta la línea 20 y descomenta la siguiente
'--------------------------------------------------------------------------------------------------------
Dim resultado As Double

10    On Error GoTo LinErr
20    resultado = DLookup("[numero]", "test", "[idtest]=5")
'   Err.Raise (94)

30    Exit Sub
    
LinErr:
    Debug.Print "Error" & Space(25 - Len("Error")) & ": " & Err.Number & vbNewLine & _
                "Descripción" & Space(25 - Len("Descripción")) & ": "; Err.Description & vbNewLine & _
                "Id de la ayuda" & Space(25 - Len("Id de la ayuda")) & ": " & Err.HelpContext & vbNewLine & _
                "Fuente" & Space(25 - Len("Fuente")) & ": " & Err.Source & vbNewLine & _
                "Error en dll" & Space(25 - Len("Error en dll")) & ": " & Err.LastDllError & vbNewLine & _
                "Mas información" & Space(25 - Len("Mas información")) & ": " & Err.HelpFile & vbNewLine & _
                "Ha ocurrido un error en" & Space(25 - Len("Ha ocurrido un error en")) & ": " & Erl

'   Err.Clear
'   Debug.print "Se ha vacíado el objeto Err"
'   Debug.Print "Error" & Space(25 - Len("Error")) & ": " & Err.Number

End Sub