Módulo de clase: clsErrorhandler
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-clsErrorHandler/
'                     Destello 324
'---------------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : clsErrorHandler
' Autor original    : Microsoft | NorthWind Traders | Developer edition
' Adaptado          : Luis Viadel | luisviadel@cowtechnologies.net
' Creado            : 2023
' Propósito         : establecer un objeto propio de control de errores que genere un log para el programador
'---------------------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/189fb41b-cc3a-4999-a6d2-ba89f72d2870
'                     Esta es una clase estática, lo que significa que el desarrollador no necesita instanciarla, siempre está ahí.
'                     Para usarlo en sus propios proyectos, debe usar este procedimiento:
'                     Exportar archivo fuera de esta base de datos e Importar archivo a su base de datos.
'                     Un simple copiar/pegar el código NO es suficiente.
'
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub clsErrorHandler_test()
'10        On Error GoTo Err_Handler
'
'          'Pon aquí el código que desees
'
'Exit_Handler:
'20        Exit Sub
'
'Err_Handler:
'30        clsErrorHandler.HandleError "TuModulo", "TuProcedimiento", err.number, err.description, Erl
'40        Resume Exit_Handler
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Const LOGFILENAME As String = "MiLogDeErrores.log"
Private m_strLogFilePath As String

Private Sub Class_Initialize()
10        On Error GoTo Err_Handler

'Crea la ruta completa al archivo de registro de errores en la misma carpeta que la base de datos actual, & _
concatenando la carpeta, una barra invertida y un nombre de archivo.
20        m_strLogFilePath = CurrentProject.Path & "\" & LOGFILENAME

Exit_Handler:
30        Exit Sub

Err_Handler:
40        clsErrorHandler.HandleError "clsErrorHandler", "Class_Initialize", Err.Number, Err.Description, Erl
50        Resume Exit_Handler
End Sub

'NOTE:
'   This main error handler should not itself have an error handler.
Public Sub HandleError(ByVal strModuleName As String, ByVal strProcedureName As String, ErrorNumber As Long, ErrDescription As String, Erl As Long)
Dim strMsg As String

10        If ErrorNumber = 2501 Then Exit Sub    '2501 = la acción de abrir el formulario ha sido cancelada. No se trata realmente de un error.

'Registra el error, incluso antes de mostrárselo al usuario.
20        LogErrorToFile ErrDescription, ErrorNumber, Erl, strModuleName, strProcedureName

'Crea el mensaje concatenando cadenas con la información
30        strMsg = "Se ha producido el siguiente error: " & vbCrLf
40        strMsg = strMsg & ErrDescription & vbCrLf
50        If Erl <> 0 Then strMsg = strMsg & "en la linea " & Erl & vbCrLf
60        strMsg = strMsg & strModuleName & "." & strProcedureName & vbCrLf & vbCrLf
70        strMsg = strMsg & "Número de error: " & ErrorNumber & vbCrLf & vbCrLf
80        strMsg = strMsg & "Por favor, notifíquelo al administrador del sistema. Gracias."

' Muestra el mensaje al usuario. Si ejecutas ACCDB puedes presionar Ctrl+Break+Step en el depurador para acceder al procedimiento de llamada.
90        MsgBox strMsg, vbCritical

End Sub

Private Sub LogErrorToFile(ByVal strError As String, ByVal lngError As Long, ByVal lngErrorLine As Long, ByVal strModuleName As String, ByVal strProcedureName As String)
10        On Error GoTo Err_Handler

          Dim intFile         As Integer

20        intFile = FreeFile
30        Open m_strLogFilePath For Append As intFile

40        Print #intFile, "Fecha y hora   : " & Now
50        Print #intFile, "Módulo         : " & strModuleName
60        Print #intFile, "Procedimiento  : " & strProcedureName
70        Print #intFile, "Descripción    : " & strError
80        Print #intFile, "Número de error: " & lngError
90        Print #intFile, "Línea          : " & IIf(lngErrorLine = 0, "<none>", lngErrorLine)
100       Print #intFile, "Usuario        : " & Environ$("USERNAME")
110       Print #intFile, "Equipo         : " & Environ$("COMPUTERNAME")
120       Print #intFile, "---------------"

130       Close #intFile

Exit_Handler:
140       Exit Sub

Err_Handler:
150       clsErrorHandler.HandleError "clsErrorHandler", "LogErrorToFile", Err.Number, Err.Description, Erl
160       Resume Exit_Handler

End Sub