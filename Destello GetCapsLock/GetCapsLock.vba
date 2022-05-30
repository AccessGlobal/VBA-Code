'Incluir en un módulo estándar
Option Compare Database
Option Explicit

Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Function GetCapslock() As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-detectar-capslock
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetCapslock
' Autor             : desconocido
' Adaptado          : Luis Viadel | https://cowtechnologies.net
' Propósito         : capturar el estado de una tecla
' Retorno           : devuelve un valor integer igual a "1" si estáactivada y "0" si está desactivada
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y abrir el
'                     formulario para ver su funcionamiento.
'
'Private Sub Form_Open(Cancel As Integer)
'
'If GetCapslock = 1 Then Debug.print "Está activada"
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

    GetCapslock = GetKeyState(vbKeyCapital)

End Function

'Incluir elos eventos del form
Private Sub contrasena_Change()

If GetCapslock = 1 Then
    Me.BloqueoTxt.Visible = True
Else
    Me.BloqueoTxt.Visible = False
End If

End Sub

Private Sub contrasena_GotFocus()

If GetCapslock = 1 Then
    Me.BloqueoTxt.Visible = True
Else
    Me.BloqueoTxt.Visible = False
End If

End Sub
