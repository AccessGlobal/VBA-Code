'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : modAlbaDom
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : establecer las 8 funciones de dominio con el mismo funcionamiento que las DFunctions de VBA (DMax, DMin, DFirst, DLast, DSum, DCount,
'                     DAvg, DLookUP) basándose en sentencias SQL
' ¿Cómo funciona?   : los argumentos de las funciones son comunes y lo son a las DFunctions (Expresión, Dominio, criterios)
'                     El mçodulo está compuesto por 8 funciones_
'                     sAvg:calcula la media de un conjunto de registros
'                     sCount: cuenta los elementos de un conjunto de registros
'                     sSum:suma los elementos de un conjunto de registros
'                     sLookup:localiza un valor concreto de un conjunto de registros
'                     sMax: localiza el valor máximo de un conjunto de registros
'                     sMin: localiza el valor mínimo de un conjunto de registros
'                     sFirst: identifica el primer registro de un conjunto de registros
'                     sLast: identifica el último registro de un conjunto de registros
' Entradas          : sCampo       Obligatorio    Valor string que representa la expresión del conjunto de datos, un campo
'                     sTabla       Obligatorio    Valor string que representa la expresión del conjunto de datos, una tabla o consulta
'                     sDonde       Opcional       Valor string que representa los criterios del conjunto de datos
' Salidas           : Valores variant o long dependiendo de la función
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'                     Este test
'
'Sub FuncionesDominioAlba_test()
'Dim resultado, ResultadoDFunctions
'Dim DblResultado As Double
'
''DAvg
'    resultado = sAvg("Numero", "test")
'    ResultadoDFunctions = DAvg("Numero", "test")
'    DblResultado = FormatNumber(resultado, 2)
'    Debug.Print "La media de Alba es" & Space(45 - Len("La media de Alba es")) & ": " & DblResultado
'    DblResultado = FormatNumber(ResultadoDFunctions, 2)
'    Debug.Print "La media de DFunctions es" & Space(45 - Len("La media de DFunctions es")) & ": " & DblResultado
'    Debug.Print vbNullString
''DCount
'    resultado = sCount("Numero", "test")
'    ResultadoDFunctions = DCount("Numero", "test")
'    Debug.Print "La tabla contiene según alba" & Space(45 - Len("La tabla contiene según Alba")) & ": " & resultado & " registros"
'    Debug.Print "La tabla contiene según DFunctions" & Space(45 - Len("La tabla contiene según DFunctions")) & ": " & ResultadoDFunctions & " registros"
'    Debug.Print vbNullString
''DSum
'    resultado = sSum("Numero", "test")
'    ResultadoDFunctions = DSum("Numero", "test")
'    Debug.Print "La suma de los registros de Alba es" & Space(45 - Len("La suma de los registros de Alba es")) & ": " & resultado
'    Debug.Print "La suma de los registros de DFunctions es" & Space(45 - Len("La suma de los registros de DFunctions es")) & ": " & resultado
'    Debug.Print vbNullString
''DLookUp
'    resultado = sLookup("Numero", "test", "idtest=2")
'    ResultadoDFunctions = DLookup("Numero", "test")
'    Debug.Print "El valor del idtest=2 según Alba es" & Space(45 - Len("El valor del idtest=2 según Alba es")) & ": " & resultado
'    Debug.Print "El valor del idtest=2 según DFunctions es" & Space(45 - Len("El valor del idtest=2 según DFunctions es")) & ": " & resultado
'    Debug.Print vbNullString
''DMax
'    resultado = sMax("Numero", "test")
'    ResultadoDFunctions = DMax("Numero", "test")
'    Debug.Print "El valor máximo según Alba es" & Space(45 - Len("El valor máximo según Alba es")) & ": " & resultado
'    Debug.Print "El valor máximo según DFunctions es" & Space(45 - Len("El valor máximo según DFunctions es")) & ": " & resultado
'    Debug.Print vbNullString
''DMin
'    resultado = sMin("Numero", "test")
'    ResultadoDFunctions = DMin("Numero", "test")
'    Debug.Print "El valor mínimo según Alba es" & Space(45 - Len("El valor mínimo según Alba es")) & ": " & resultado
'    Debug.Print "El valor mínimo según DFunctions es" & Space(45 - Len("El valor mínimo según DFunctions es")) & ": " & resultado
'    Debug.Print vbNullString
''DFirst
'    resultado = sFirst("idtest", "test")
'    ResultadoDFunctions = DFirst("Numero", "test")
'    Debug.Print "El primer id de registro según Alba es" & Space(45 - Len("El primer id de registro según Alba es")) & ": " & resultado
'    Debug.Print "El primer id de registro según DFunctions es" & Space(45 - Len("El primer id de registro según DFunctions es")) & ": " & resultado
'    Debug.Print vbNullString
''DLast
'    resultado = sLast("idtest", "test")
'    ResultadoDFunctions = DLast("Numero", "test")
'    Debug.Print "El último id de registro según Alba es" & Space(45 - Len("El último id de registro según Alba es")) & ": " & resultado
'    Debug.Print "El último id de registro según DFunctions es" & Space(45 - Len("El último id de registro según DFunctions es")) & ": " & resultado
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

Function sAvg(sCampo As String, sTabla As String, Optional sDonde As String) As Variant
'----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : SAvg
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve la media de valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : Variant
'----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub sAvg_test()
' Dim media
'
'   media = sAvg("campo", "tabla")
'   Debug.Print "La media es: " & media
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sAvg_Error
    
    MiSQL = "SELECT AVG(" & wzBracketString(sCampo, 0) & ") as Media FROM " & wzBracketString(sTabla, 0)
    If Trim(sDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

        If Rs.BOF And Rs.EOF Then
            sAvg = Null
        Else
            Rs.MoveFirst
            sAvg = Rs!media
        End If
    
        Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sAvg_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sAvg del Módulo modAlbaDom"

End Function

Function sCount(sCampo As String, sTabla As String, Optional sDonde As String = "") As Long
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sCount
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve la cuenta de registros de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Long
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sCount_Error

    MiSQL = "SELECT COUNT(" & wzBracketString(sCampo, 1) & ") as MiCount FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sCount = 0
    Else
        Rs.MoveFirst
        sCount = Rs!MiCount
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sCount_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sCount del Módulo modAlbaDom"

End Function

Function sFirst(sCampo As String, sTabla As String, Optional sDonde As String) As Variant
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sFirst
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el primer valor de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Long
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sFirst_Error

    MiSQL = "SELECT FIRST(" & wzBracketString(sCampo, 1) & ") as MiFirst FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sFirst = Null
    Else
        Rs.MoveFirst
        sFirst = Rs!MiFirst
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sFirst_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sFirst del Módulo modAlbaDom"
    Resume
End Function

Function sLast(sCampo As String, sTabla As String, Optional sDonde As String) As Variant
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sLast
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el último valor de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Variant
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sLast_Error

    MiSQL = "SELECT LAST(" & wzBracketString(sCampo, 1) & ") as MiLast FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sLast = Null
    Else
        Rs.MoveFirst
        sLast = Rs!MiLast
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sLast_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sLast del Módulo modAlbaDom"

End Function

Function sLookup(sCampo As String, sTabla As String, Optional sDonde As String = "") As Variant
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sLookup
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el primer valor encontrado de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Variant
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sLookup_Error

    MiSQL = "SELECT " & wzBracketString(sCampo, 1) & " FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sLookup = Null
    Else
        Rs.MoveFirst
        sLookup = Rs.Fields(wzBracketString(sCampo, 1)).value
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sLookup_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sLookup del Módulo modAlbaDom"
    Resume Next
End Function

Function sMax(sCampo As String, sTabla As String, Optional sDonde As String) As Variant
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sMax
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el máximo de los valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Variant
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String


    On Error GoTo sMax_Error

    MiSQL = "SELECT MAX(" & wzBracketString(sCampo, 1) & ") as MiMax FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sMax = Null
    Else
        Rs.MoveFirst
        sMax = Rs!MiMax
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sMax_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sMax del Módulo modAlbaDom"

End Function

Function sMin(sCampo As String, sTabla As String, Optional sDonde As String) As Variant
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sMin
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el mínimo de los valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Variant
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sMin_Error

    MiSQL = "SELECT MIN(" & wzBracketString(sCampo, 1) & ") as MiMin FROM " & wzBracketString(sTabla, 1)
    If Trim(sDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

    If Rs.BOF And Rs.EOF Then
        sMin = Null
    Else
        Rs.MoveFirst
        sMin = Rs!MiMin
    End If

    Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sMin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sMin del Módulo modAlbaDom"

End Function

Function sSum(sCampo As String, sTabla As String, Optional sDonde As String) As Long
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funciones-de-dominio-de-alba/
'--------------------------------------------------------------------------------------------------------
' Título            : sSum
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve el mínimo de los valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'-------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'--------------------------------------------------------------------------------------------------------
' Retorno           : Long
'--------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String

    On Error GoTo sSum_Error

    MiSQL = "SELECT SUM(" & wzBracketString(sCampo, 1) & ") as Suma FROM " & wzBracketString(sTabla, 1)
    
    If Trim(sDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & sDonde
    End If

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

        If Rs.BOF And Rs.EOF Then
            sSum = 0
        Else
            Rs.MoveFirst
            sSum = Rs!Suma
        End If
    
        Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sSum_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sSum del Módulo modAlbaDom"

End Function

