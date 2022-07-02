Option Compare Database
Option Explicit

Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, _
                                                      ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                                                      ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long

Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Sub WipeEffect(frm As Form, lngOpt As Long, lngIncrement As Long)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-efectos-wipe-y-shrink
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : WipeEffect
' Autor             : Candace Tripp | http://www.candace-tripp.net/
' Fecha             : anterior al 9 de mayo de 2005
' Propósito         : provocar el efecto wipe en formularios de Access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://www.access-programmers.co.uk/forums/threads/wipe-effects.86382/
'                     https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-movewindow
'                     https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en el evento de cierre del formulario que desees.
'
' Private Sub Form_Close()
' Dim lngIncrement As Long
'
'    lngIncrement = 100
'    Call WipeEffect(Me, 1, lngIncrement)
'
' End Sub
'
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim r As RECT
Dim lngRet As Long
Dim lngX As Long
Dim lngTop As Long
Dim lngLeft As Long
Dim factor As Long

Dim lngFormHeight As Long
Dim lngFormWidth As Long
    
Dim lngIncrementW As Long
Dim lngIncrementH As Long
    
        
    lngRet = GetWindowRect(frm.hwnd, r)
    lngFormWidth = r.right - r.left
    lngFormHeight = r.bottom - r.top
        
    lngIncrementW = lngFormWidth \ lngIncrement
    lngIncrementH = lngFormHeight \ lngIncrement
    
    Select Case lngOpt
        Case 1 ' wipe up
            For lngX = 1 To lngIncrement
                lngRet = MoveWindow(frm.hwnd, r.left, r.top, _
                        lngFormWidth, lngFormHeight - lngX * lngIncrementH, 1)
            Next lngX
        
        Case 2 ' wipe down
            For lngX = 1 To lngIncrement
                lngRet = MoveWindow(frm.hwnd, r.left, r.top + lngX * lngIncrementH, _
                        lngFormWidth, lngFormHeight - lngX * lngIncrementH, 1)
            Next lngX
        
        Case 3 ' wipe right
            For lngX = 1 To lngIncrement
                lngRet = MoveWindow(frm.hwnd, r.left + lngX * lngIncrementW, r.top, _
                        lngFormWidth - lngX * lngIncrementW, lngFormHeight, 1)
            Next lngX
        
        Case 4 ' wipe left
            For lngX = 1 To lngIncrement
                lngRet = MoveWindow(frm.hwnd, r.left, r.top, _
                        lngFormWidth - lngX * lngIncrementW, lngFormHeight, 1)
            Next lngX
        
        Case 5 ' shrink/move
            For lngX = 1 To lngIncrement
                lngRet = MoveWindow(frm.hwnd, r.left - lngX * lngIncrementW, _
                         r.top + lngX * lngIncrementH, _
                         lngFormWidth - lngX * lngIncrementW, _
                         lngFormHeight - lngX * lngIncrementH, 1)
            Next lngX
    
    Case Else ' shiver
        factor = 30
        For lngX = 1 To 2500
            If lngX Mod 4 = 0 Then
                lngLeft = r.left - factor
                lngTop = r.top - factor
            ElseIf lngX Mod 3 = 0 Then
                lngLeft = r.left - factor
                lngTop = r.top + factor
            ElseIf lngX Mod 2 = 0 Then
                lngLeft = r.left + factor
                lngTop = r.top - factor
            Else
                lngLeft = r.left + factor
                lngTop = r.top + factor
            End If
            lngRet = MoveWindow(frm.hwnd, _
                     lngLeft, _
                     lngTop, _
                     lngFormWidth, _
                     lngFormHeight, 1)
        Next lngX
        MsgBox "Brrrrrrrr!!  I think I hab a code.", vbCritical, "Code"
    End Select
    

End Sub

