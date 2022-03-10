Public Declare PtrSafe Sub ReleaseCapture Lib "user32" ()
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Function MoverForm(Form1 As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-mover-formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Mover un formulario
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Propósito         : Poder mover un formulario con el ratón emulando la propiedad mover formulario de Windows
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Form1      Obligatorio/Opcional      Especifica el formulario que estamos moviendo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : Windows API
' Más información   : https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-releasecapture
'                     https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub MoverForm_test()
'
'    Call MoverForm(Me)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
ReleaseCapture
Call SendMessage(Form1.hwnd, &HA1, 2, 0&)
End Function