'Módulo estándar: "ModAggDomRem"
Option Compare Database
Option Explicit

Public Db As Object
Public rst As Object
Public MiSQL As String
Public Valor As Variant

Public Function rCount(ByVal rCampo As String, ByVal rTabla As String, Optional rDonde As String = "", Optional dbPath As String, Optional UseJetLink As Boolean = True) As Long
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rCount
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función que devuelve la cuenta de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor long con el número de registros
'--------------------------------------------------------------------------------------------------
    
    rCount = 0
    On Error GoTo lbError

    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT COUNT(" & rCampo & ") as MiValor FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

        If rst.BOF And rst.EOF Then
            Valor = 0
        Else
            rst.MoveFirst
            Valor = rst!MiValor
        End If
    
    rCount = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rCount del Módulo ModAggDomRem"
    
lbFinally:
    On Error Resume Next
    
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function

Function rLookUp(ByVal rCampo As String, ByVal rTabla As String, Optional rDonde As String = "", Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rLookUp
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para buscar un dato en una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------
    
    On Error GoTo lbError

    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT " & rCampo & " FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

        If rst.BOF And rst.EOF Then
            Valor = Null
        Else
            rst.MoveFirst
            Valor = rst.Fields(rCampo).Value
        End If

    rLookUp = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rLookUp del Módulo ModAggDomRem"

lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function

Public Function rMax(ByVal rCampo As String, ByVal rTabla As String, Optional rDonde As String = "", Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rMax
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para buscar el máximo de un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------
    
    rMax = Null
    On Error GoTo lbError

    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT MAX(" & rCampo & ") as MiValor FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = rst!MiValor
    End If

    rMax = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rMax del Módulo ModAggDomRem"
    
lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function

Public Function rMin(ByVal rCampo As String, ByVal rTabla As String, Optional rDonde As String = "", Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rMin
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para buscar el mínimo de un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------
        
    rMin = Null
    On Error GoTo lbError

    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT MIN(" & rCampo & ") as MiValor FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = rst!MiValor
    End If

    rMin = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rMin del Módulo ModAggDomRem"
    
lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function

Public Function rSum(ByVal rCampo As String, ByVal rTabla As String, Optional rDonde As String, Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rSum
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para sumar un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------
    
    rSum = Null
    On Error GoTo lbError

    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT SUM(" & rCampo & ") as MiValor FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = Nz(rst!MiValor, 0)
    End If

    rSum = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rSum del Módulo ModAggDomRem"
       
lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function

Function rAvg(rCampo As String, rTabla As String, Optional rDonde As String, Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rAvg
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para obtener la media un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------

    On Error GoTo lbError
    rAvg = Null
    
    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT AVG(" & rCampo & ") as Media FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde) & "" <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = Nz(rst!MiValor, 0)
    End If
    
    rAvg = Valor
    GoTo lbFinally

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rAvg del Módulo ModAggDomRem"

lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0

End Function


Function rFirst(rCampo As String, rTabla As String, Optional rDonde As String, Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rFirst
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para obtener el primer registro de un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------

    On Error GoTo lbError

    rFirst = Null
    
    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT FIRST(" & rCampo & ") as MiFirst FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)

    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = Nz(rst!MiValor, 0)
    End If
    

    rFirst = Valor
    GoTo lbFinally
    

lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally
    
lbError:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rFirst del Módulo ModAggDomRem"
  
lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0
  
End Function

Function rLast(rCampo As String, rTabla As String, Optional rDonde As String, Optional dbPath As String, Optional UseJetLink As Boolean = True) As Variant
'--------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funciones-de-dominio-de-alba-funciones-remotas
'--------------------------------------------------------------------------------------------------
' Título            : rLast
' Autor original    : Alba Salvá
' Creado            : 12/03/2023
' Propósito         : función para obtener el último registro de un conjunto de registros de una BD remota
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable        Modo          Descripción
'--------------------------------------------------------------------------------------------------
'                     rCampo        Obligatorio   Nombre del campo
'                     rTabla        Obligatorio   Nombre de la tabla
'                     rDonde        Opcional      Criterios adicionales para la búsqueda
'                     dbPath        Opcional      Para buscar en bases de datos externas
'                     UseJetLink    Opcional      Para buscar usando JetLink (solo BBDD Access)
'--------------------------------------------------------------------------------------------------
' Retorno           : valor variant con el dato buscado
'--------------------------------------------------------------------------------------------------
    On Error GoTo lbError

    rLast = Null
    
    If dbPath = "" Or (dbPath <> "" And UseJetLink) Then
        Set Db = CurrentDb
    Else
        Set Db = DBEngine.OpenDatabase(dbPath)
    End If

    If Not wzBracketString(rCampo, 1) Then GoTo lbErrorBracketString
    If Not wzBracketString(rTabla, 1) Then GoTo lbErrorBracketString

    MiSQL = "SELECT LAST(" & wzBracketString(rCampo, 1) & ") as MiLast FROM "
    
    If UseJetLink And dbPath <> "" Then
        MiSQL = MiSQL & "[" & dbPath & "]."
    End If
    MiSQL = MiSQL & rTabla
    
    If Trim(rDonde & "") <> "" Then
        MiSQL = MiSQL & " WHERE " & rDonde
    End If

    Set rst = Db.OpenRecordset(MiSQL, dbOpenSnapshot, dbReadOnly)
    
    If rst.BOF And rst.EOF Then
        Valor = Null
    Else
        rst.MoveFirst
        Valor = rst!MiLast
    End If

    rLast = Valor
    GoTo lbFinally
    
lbErrorBracketString:

    MsgBox "Error en conversión BracketString en Function rCount del Módulo ModAggDomRem"
    
    GoTo lbFinally

lbError:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function rLast del Módulo ModAggDomRem"
    
lbFinally:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    If Not Db Is Nothing Then Db.Close
    Set Db = Nothing

    On Error GoTo 0
  
End Function
