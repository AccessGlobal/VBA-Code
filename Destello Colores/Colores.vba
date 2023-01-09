Option Compare Database
Option Explicit

Private Declare PtrSafe Sub wlib_AccColorDialog Lib "msaccess.exe" Alias "#53" (ByVal hwnd As Long, lngRGB As Long)

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-descubre-los-colores
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Form_Color
' Autor original    : Varios autores desconocidos
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : en algún momento de 2011
' Propósito         : obtener el código de colores en long, hex y RGB
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/rgb-function
'                     https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/hex-function
'                     https://stackoverflow.com/questions/10093374/what-does-ired-icolor-mod-256-mean
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub ColorIn_Click()
Dim col As Long
Dim r As Long, g As Long, b As Long
Dim hex1 As String, hex2 As String, hex3 As String

'Mostramos el color por defecto
    col = 2843349
'Llamamos a la API para obtener el selector de color
    wlib_AccColorDialog Screen.ActiveForm.hwnd, col

    Me.pcolorint = col
    
'Cambia el color de los controles para el efecto camiseta
    Me.ColorPrueba.BackColor = col
    Me.ColorPrueba.ForeColor = col
    Me.ColorPrueba = col
    
'Calcula RGB
    r = col Mod 256
    g = (col \ 256) Mod 256
    b = (col \ 256 \ 256) Mod 256
    
    Me.pcolorrgb = "RGB(" & r & "," & g & "," & b & ")"
    
'Calcula Hexadecimal
    hex1 = hex(r)
    hex2 = hex(g)
    hex3 = hex(b)
       
    Me.pcolorhex = "#" & hex1 & hex2 & hex3

End Sub
