'Formulario, botones de comando
Private Declare PtrSafe Sub wlib_AccColorDialog Lib "msaccess.exe" Alias "#53" (ByVal hwnd As Long, lngRGB As Long)

Private Sub btnCambio_Click()
'Cambia el color de un treView a un color predefinido

    Call SetTVBackColour(color1, Me.TV1)
    Call SetTVBackColour(color2, Me.TV2)
    Call SetTVBackColour(color3, Me.TV3)
      
End Sub

Private Sub btnSeleccion_Click()
Dim col As Long

'Llamamos a la API para obtener el selector de color
    wlib_AccColorDialog Screen.ActiveForm.hwnd, col

    Call SetTVBackColour(col, Me.TV2)

End Sub


'Módulo estándar: "modTV"
Option Compare Database
Option Explicit

Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const TVS_HASLINES As Long = 2
Public Const TVM_SETBKCOLOR As Long = (&H1100 + 29)

Public Const color1 = 5389869
Public Const color2 = 12881497
Public Const color3 = 15128261

Public Function SetTVBackColour(clrref As Long, TV1 As Object)
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-cambiar-color-de-treeview/
'                     Destello formativo 373
' Fuente original   : http://vbnet.mvps.org/index.html?code/comctl/tveffects.htm
'--------------------------------------------------------------------------------------------------------
' Título            : SetTVBackColour
' Autor original    : VBnet - Randy Birch
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : Friday April 17, 1998
' Modificado        : Monday December 26, 2011
' Adaptado          : Luis Viadel - 25/05/2021
' Propósito         : cambiar el color de un objeto TreeView
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA.
'
'Sub SetTVBackColour_Test()
'Dim BackColor As Long
'Dim ctrl As control
'
'   Call SetTVBackColour(BackColor, ctrl1)
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim hwndTV As Long, style As Long
   
    hwndTV = TV1.hwnd
   
'Change the background
    Call SendMessage(hwndTV, TVM_SETBKCOLOR, 0, ByVal clrref)
   
'reset the treeview style so the tree lines appear properly
    style = GetWindowLong(TV1.hwnd, GWL_STYLE)
   
'if the treeview has lines, temporarily remove them so the back repaints to the
'selected colour, then restore
    If style And TVS_HASLINES Then
        Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
        Call SetWindowLong(hwndTV, GWL_STYLE, style)
    End If
  
End Function

