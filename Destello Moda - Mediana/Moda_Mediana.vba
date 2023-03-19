Function sMedian(sCampo As String, sTabla As String, Optional sDonde As String) As Double
'----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-moda-mediana-y-algo-mas
'----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : sMedian
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve la mediana de los valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : Double
'----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub sMedian_test()
' Dim mediana
'
'   mediana = sMedian("campo", "tabla")
'   Debug.Print "La mediana es: " & mediana
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim sSql As String
Dim NumReg As Long
Dim Valor1 As Double
Dim sTbl As String
Dim sFld As String
    
    sTbl = wzBracketString(sTabla, 1)
    sFld = wzBracketString(sCampo, 1)
    
    sMedian = False
    
    On Error GoTo sMedian_Error
    
    sSql = "SELECT " & sFld & " FROM " & sTbl & " ORDER BY " & sFld
    
    If Trim(sDonde) & "" <> "" Then
        sSql = sSql & " WHERE " & sDonde
    End If
    
    Set Rs = CurrentDb.OpenRecordset(sSql, dbOpenDynaset)
        If Rs.RecordCount > 0 Then
            Rs.MoveLast
            NumReg = Rs.RecordCount
            
            If NumReg Mod 2 = 1 Then 'Es impar.
                Rs.MoveFirst
                Rs.Move NumReg / 2
                sMedian = Rs(sFld)
            Else ' Es par.
                Rs.MoveFirst
                Rs.Move NumReg / 2
                Valor1 = Rs(sFld)
                Rs.MovePrevious
                sMedian = (Valor1 + Rs(sFld)) / 2
            End If
        End If
        
        Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sMedian_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sMedian del Módulo modAggDom"

End Function

Function sModa(sCampo As String, sTabla As String, Optional sDonde As String) As Double
'----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-moda-mediana-y-algo-mas
'----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : sModa
' Autor original    : Alba Salvá
' Creado            : 06/12/2010
' Propósito         : Función que devuelve la moda de los valores de una tabla
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------
'                     sCampo           Obligatorio    Nombre del campo
'                     sTabla           Obligatorio    Nombre de la tabla
'                     sDonde           Opcional       Criterios adicionales para la búsqueda
'----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : Dble
'----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub sModa_test()
' Dim moda
'
'   moda = sModa("campo", "tabla")
'   Debug.Print "La moda es: " & moda
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Rs As Recordset
Dim MiSQL As String
Dim sTbl As String
Dim sFld As String
    
    sTbl = wzBracketString(sTabla, 1)
    sFld = wzBracketString(sCampo, 1)
    

    On Error GoTo sModa_Error

    MiSQL = "SELECT " & sFld & vbCrLf & _
           " FROM (SELECT " & sFld & ", Count(" & sFld & ") AS Frecuencia FROM " & sTbl & " GROUP BY " & sFld
           
    If Trim(sDonde) & "" <> "" Then
        MiSQL = MiSQL & " HAVING " & sDonde
    End If
           
    MiSQL = MiSQL & ") AS Datos_Frecuencia " & vbCrLf & _
           "WHERE frecuencia = (SELECT MAX(Frecuencia) FROM (SELECT numero, Count(numero) AS Frecuencia FROM test GROUP BY numero));"

    Set Rs = CurrentDb.OpenRecordset(MiSQL)

        If Rs.BOF And Rs.EOF Then
            sModa = Null
        Else
            Rs.MoveFirst
            sModa = Rs(sFld)
        End If
    
        Rs.Close
    Set Rs = Nothing

    On Error GoTo 0
    Exit Function

sModa_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en Function sModa del Módulo modAlbaStats"

End Function