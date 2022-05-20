Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-formularios-redondos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : modRoundForm
' Autor             : Luis Viadel | https://cowtechnologies.net
' Fecha             : febrero 2021
' Propósito         : cambiar el aspecto de un formulario, clásicamente rectangular, por un diseño elíptico. En este caso será circular porque
'                     partimos de un formulario cuadrado
' Retorno           : Sin retorno
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowrgn
'                     https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createellipticrgn
'                     https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdc
'                     https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdevicecaps
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test, que tiene que estar incluido en el
'                     evento "Al abrir" del formulario que queramos modificar.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y abrir el
'                     formulario para ver su funcionamiento.
'
'Private Sub Form_Open(Cancel As Integer)
'Dim Xs As Long, Ys As Long
'
'Xs = Me.Width / TwipsPerPixelX(Me)
'Ys = Me.InsideHeight / TwipsPerPixelY()
'
'SetWindowRgn hwnd, CreateEllipticRgn(0, 0, Xs, Ys), True
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
                ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal Index As Long) As Long

Const HWND_DESKTOP As Long = 0
Const LOGPIXELSX As Long = 88
Const LOGPIXELSY As Long = 90

Function TwipsPerPixelX(frm As Form) As Single
Dim lngDC As Long

lngDC = GetDC(frm.hwnd)
'Nos devuelve el valor en pixels por pulgada
TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX) 'Traspasamos el valor a Twips

ReleaseDC frm.hwnd, lngDC

End Function

Function TwipsPerPixelY() As Single
Dim lngDC As Long

lngDC = GetDC(HWND_DESKTOP)

TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)

ReleaseDC HWND_DESKTOP, lngDC

End Function
