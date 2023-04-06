Public Function ObtenerCode(ByVal sModuleName As String, ByVal sProcName As String, Optional bInclCabecera As Boolean = True) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-obtener-código-concreto
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ObtenerCode
' Autor original    : Alba Salvá
' Creado            : marzo 2023
' Propósito         : Obtener todo el código de un procedimiento
' Retorno           : cadena con el código obtenido
' Argumento/s       : La sintaxis del procedimiento o función consta de los siguientes argumentos:
'                     Parte                      Modo                    Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     sModuleName               Obligatorio          Nombre del módulo que contiene el procedimiento a buscar
'                     sProcName                 Obligatorio          Nombre del procedimiento a extraer el textor
'                     bInclCabecera             Obligatorio          True/False - Indica si se incluye la cabecera del procedimiento en la
'                                                                    salida del texto
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este test. Copia el bloque siguiente al
'                     portapapeles y pégalo en el editor de VBA en el evento de tu elección.
'                     Descomenta las líneas y pulsa F5 para ver su funcionamiento.
'
' ? ObtenerCode("Module1", "fOSUserName")
' ? ObtenerCode("Module1", "fOSUserName", False)
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim oModule               As Object  'CodeModule
Dim lProcStart            As Long
Dim lProcBodyStart        As Long
Dim lProcNoLines          As Long
    
    'Const vbext_pk_Proc = 0 'Requerido en caso de usar Late Binding
 
 
    On Error GoTo lbError
 
    Set oModule = Application.VBE.ActiveVBProject.VBComponents(sModuleName).CodeModule ' Establecemos el objeto módulo y su código
        lProcStart = oModule.ProcStartLine(sProcName, vbext_pk_Proc) ' Obtenemos el principio del procedimiento
        lProcBodyStart = oModule.ProcBodyLine(sProcName, vbext_pk_Proc) ' Obtenemos el principio del cuerpo del procedimiento
        lProcNoLines = oModule.ProcCountLines(sProcName, vbext_pk_Proc) ' Obtenemos la longitud del procedimiento
        If bInclCabecera = True Then                                    ' Si se incluye la cabecera, tomamos todo el procedimiento
            ObtenerCode = oModule.Lines(lProcStart, lProcNoLines)
        Else                                                            ' en caso contrario, sólo el cuerpo
            lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
            ObtenerCode = oModule.Lines(lProcBodyStart, lProcNoLines)
        End If
 
    GoTo lbFinally
 
lbError:
    'Se produce el error 35 si el procedimiento no se encuentra
    MsgBox "Ha ocurrido un error" & vbCrLf & vbCrLf & _
           "Código de Error : " & Err.Number & vbCrLf & _
           "Origen del Error : ObtenerCode" & vbCrLf & _
           "Descripción Error : " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Línea No: " & Erl) _
           , vbOKOnly + vbCritical, "¡Ha ocurrido un error"

lbFinally:
    On Error Resume Next
    If Not oModule Is Nothing Then Set oModule = Nothing

End Function

