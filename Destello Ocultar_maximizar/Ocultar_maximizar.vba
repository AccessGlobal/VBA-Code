'Declarar a nivel de módulo
Public Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4
Const wFlags = SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED + SWP_NOMOVE
Const FLAGS_COMBI = WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_SYSMENU
Function BotonMaxVisible(bEnable As Boolean) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/ocultar-maximizar-access/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : BotonMaxVisible
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Actualizado       : abril 2011
' Propósito         : bloquea la opción maximizar la ventana de Access
' Retorno           : long
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte                Modo        Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     activado          Obligatorio    El valor Boolean especifica si el botón está bloqueado o no
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub BotonMaxVisible_test()
'
'   BotonMaxVisible true
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim hwnd As Long
Dim nIndex As Long
Dim dwNewLong As Long
Dim dwLong As Long
hwnd = hWndAccessApp
nIndex = GWL_STYLE
dwLong = GetWindowLong(hwnd, nIndex)
If bEnable Then
dwNewLong = (dwLong Or WS_MAXIMIZEBOX)
Else
dwNewLong = (dwLong And Not WS_MAXIMIZEBOX)
End If
Call SetWindowLong(hwnd, nIndex, dwNewLong)
Call SetWindowPos(hwnd, 0&, 0&, 0&, 0&, 0&, wFlags)
End Function