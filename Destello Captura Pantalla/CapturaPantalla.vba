'El código está diseñado para trabajar desde un módulo de formulario con dos botones

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-captura-de-pantalla
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CapturaPantalla
' Autor original    : Luis Viadel
' Propósito         : realizar una captura de pantalla y colocarla en el portapapeles
' Información       :
'                     https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-keybd_event
'                     https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes'-----------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btn1_Click()

    keybd_event 44, 0, 0&, 0&

End Sub

Private Sub btn2_Click()
    
    keybd_event 44, 1, 0&, 0&

End Sub



