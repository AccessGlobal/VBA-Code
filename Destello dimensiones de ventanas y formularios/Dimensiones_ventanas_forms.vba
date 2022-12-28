Option Compare Database
Option Explicit

Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Sub DimensionesMonitor()

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-dimensiones-de-ventanas-y-formularios
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : DimensionesMonitor
' Autor             : Luis Viadel
' Fecha             : en algún momento de 2013
' Propósito         : conocer ubicación y dimensiones de ventanas y formularios
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
'                     https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdesktopwindow
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en el evento de apertura o carga del formulario que desees.
'
' Private Sub Form_load()
' Dim rec As RECT
' Dim anchoform As Long, altoform As Long
'
'
'    Call GetWindowRect(Me.hwnd, rec)
'    anchoform=rec.Right - rec.Left
'    altoform = rec.Bottom - rec.Top
'    Debug.Print anchoform & " x " & altoform
'
' End Sub
'
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub DimensionesMonitor()
Dim rec As RECT
Dim anchomonitor As Long, altomonitor As Long

    Call GetWindowRect(GetDesktopWindow, rec)
    anchomonitor = rec.Right - rec.Left
    altomonitor = rec.Bottom - rec.Top
    Debug.Print anchomonitor & " x " & altomonitor
    
End Sub