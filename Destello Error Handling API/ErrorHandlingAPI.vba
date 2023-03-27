Option Compare Database
Option Explicit

'Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
'Const SUBLANG_DEFAULT = &H1
'Tipo de error
Const ERROR_BAD_TOKEN_TYPE = 1349&

Const SEM_FAILCRITICALERRORS = &H1
Const SEM_NOGPFAULTERRORBOX = &H2
Const SEM_NOALIGNMENTFAULTEXCEPT = &H3
Const SEM_NOOPENFILEERRORBOX = &H8000

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function GetErrorMode Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Sub ErrorHandling_test()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/error-handling-api
'                     Destello formativo 299
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ErrorHandling_test
' Autor             : Luis Viadel
' Fecha             : marzo 2023
' Propósito         : no se trata de una función ni un procedimiento entendido como tal. Es un ejemplo para mostrar el comportamiento de
'                     ciertas funciones de la API de Windows relacionadas con el tratamiento de errores de sistema
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-formatmessage
'                     https://learn.microsoft.com/en-us/windows/win32/api/errhandlingapi/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : esta función constituye en si un test. Descomenta las líneas que precises para probar las funciones.
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim strUser As String
Dim ErrorMode As Long
Dim LastError As Long
Dim LastErrorTxt As String

'Establecemos una cadena vacía de longitud=100, porque la necesita la
'función FormatMessage
    strUser = Space(100)
    
'Establecemos el modo 1 donde el sistema no muestra el cuadro de mensaje critical-error-handler.
'En su lugar, el sistema envía el error al proceso de llamada
    SetErrorMode SEM_FAILCRITICALERRORS 'mode 1
'Otros modos que podemos utilizar
'    SetErrorMode SEM_NOGPFAULTERRORBOX 'mode 2
'    SetErrorMode SEM_NOALIGNMENTFAULTEXCEPT 'mode 3
        
'Generamos un error
    SetLastError ERROR_BAD_TOKEN_TYPE
    
'Comprobamos el modo error y si es <> 1 (SEM_FAILCRITICALERR),
'porque el sistema no muestra el cuadro de mensaje critical-error-handler.
'En su lugar, el sistema envía el error al proceso de llamada.
'Lo cambiamos mediante SetError Mode
    ErrorMode = GetErrorMode
    
    If ErrorMode <> 1 Then
        SetErrorMode SEM_FAILCRITICALERRORS 'mode 1
    End If
    
'Con el nuevo modo, que aunque muestra los errores, al cambiar el modo, hemos perdido el error
'Volvemos a generar un error
    SetLastError ERROR_BAD_TOKEN_TYPE
    ErrorMode = GetErrorMode

    Debug.Print ErrorMode
    Debug.Print Err.LastDllError
    
'Para poder obtener el error mediante GetLastError tenemos que hacerlo con FormatMessage para
'poder conocer el error que se ha producido
    LastError = GetLastError
    Debug.Print LastError
    
'Utilizamos la función GetSystemErrorMessageText para obtener el texto del mensaje
    LastErrorTxt = GetSystemErrorMessageText(Err.LastDllError)
    Debug.Print LastErrorTxt
    
'En Visual Basic se recomienda no utilizar GetLastError, pero en caso de utilizarlo, esta podría ser
'una forma
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strUser, 100, ByVal 0&
    MsgBox strUser

End Sub

'Establecer este código en otro módulo estándar
Option Compare Database
Option Explicit

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0 ' from VC++ ERRORS.H file

Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long

Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/error-handling-api
'                     Destello formativo 299
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetSystemErrorMessageText
' Autor             : Chp Pearson, www.cpearson.com, chip@cpearson.com
' Fecha             : Desconocida
' Propósito         : Esta función obtiene el texto del mensaje de error del sistema que corresponde al parámetro del código de error
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Cómo funciona     : Este valor es el valor devuelto por Err.LastDLLError o por GetLastError, u ocasionalmente como el resultado devuelto
'                     por una función API de Windows.
'                     Estos NO son los números de error devueltos por Err.Number (para estos errores, use Err.Description para obtener la
'                     descripción del error). En general, debe usar Err.LastDllError en lugar de GetLastError porque, en algunas circunstancias,
'                     el valor de GetLastError se restablecerá a 0 antes de que el valor se devuelva a VBA. Err.LastDllError siempre devolverá
'                     de manera confiable el último número de error generado en una función API. La función devuelve vbNullString si se produjo
'                     un error o si no hay texto de error para el número de error especificado.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Argumentos        : La sintaxis de la función consta de un único argumento
'                     Variable          Modo          Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     ErrorNumber       Obligatorio   Número de error de sistema
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : string con la descripción del error que pasamos como parámetro
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomenta las líneas y pulsa F5 para ver su funcionamiento.
'
'Sub GetSystemErrorMessageText_test()
'Dim ErrorSistema As String
'
'    ErrorSistema = GetSystemErrorMessageText(2404&)
'
'    Debug.Print ErrorSistema
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

' Initialize the variables
    LangID = 0&   ' Default language
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
    TextLen = FORMAT_MESSAGE_TEXT_LEN

' Call FormatMessage to get the text of the error message text
' associated with ErrorNumber.
FormatMessageResult = FormatMessage( _
                        dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS, _
                        lpSource:=0&, _
                        dwMessageId:=ErrorNumber, _
                        dwLanguageId:=LangID, _
                        lpBuffer:=ErrorText, _
                        nSize:=TextLen, _
                        Arguments:=0&)

    If FormatMessageResult = 0& Then
' An error occured. Display the error number, but don't call GetSystemErrorMessageText to get the
' text, which would likely cause the error again getting us into a loop.
            MsgBox "An error occurred with the FormatMessage" & _
                   " API function call." & vbCrLf & _
                   "Error: " & CStr(Err.LastDllError) & _
                   " Hex(" & Hex(Err.LastDllError) & ")."
            GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
                "FormatMessage API function: " & CStr(Err.LastDllError) & ". No futher information" & vbCrLf & _
                "is available."
            Exit Sub
    End If
' If FormatMessageResult is not zero, it is the number of characters placed in the ErrorText variable.
' Take the left FormatMessageResult characters and return that text.
    ErrorText = Left$(ErrorText, FormatMessageResult)

' Get rid of the trailing vbCrLf, if present.
    If Len(ErrorText) >= 2 Then
        If Right$(ErrorText, 2) = vbCrLf Then
            ErrorText = Left$(ErrorText, Len(ErrorText) - 2)
        End If
    End If

' Return the error text as the result.
    GetSystemErrorMessageText = ErrorText

End Sub



