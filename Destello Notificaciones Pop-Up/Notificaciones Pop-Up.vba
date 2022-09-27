'En los eventos del formulario
Private Sub Form_Open(Cancel As Integer)

'Primero: cargamos el literal de la etiqueta de cabecera
    Me.MensaCab = "Aviso de MiApp"

'Movemos el formulario a la posición deseada
    Notificacion Me, 4, AnchoMonitor - (Me.Width / 13), AltoMonitor + TamañoBarra + 7 - (Me.Detalle.Height / 6.3), Me.Detalle.Height / 13, Me.Width / 15

'Destello formativo 82
'    Sonido ("Mi sonido")

End Sub

Sub Form_Timer()

    DoCmd.Close acForm, "MensAvisoPopUp"

End Sub

'En un módulo estandar

Option Compare Database
Option Explicit

'Módulo notificaciones
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-notificaciones-popup
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Módulo notificaciones
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : febrero 2015
' Propósito         : colocar un formulario en la parte inferior izquierda, encima de la barra de tareas de Windows
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
'                   : https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
'                   : https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdesktopwindow
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, seleccionar una fuente que no se encuentre
'                     en el sistema y pulsar F5 para ver su funcionamiento.
'
'Sub AvisoPopUp_test()
'Dim str As String
'
'    str = "Información personalizada"
'
'    Call AvisoPopUp(str)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long

Public buffRECT As RECT

Sub AvisoPopUp_test()
Dim str As String

    str = "Información personalizada"
    
    Call AvisoPopUp(str)

 End Sub

Public Function AvisoPopUp(ByVal liter As String)

    DoCmd.OpenForm "MensAvisoPopUp"

    Form_MensAvisoPopUp.Lite1 = liter
    

End Function

Function Notificacion(frm As Form, TimeSegundos As Long, cx As Long, cy As Long, cHeight As Long, cWidth As Long)
Dim HMen As Long

    HMen = frm.hwnd
    
    frm.TimerInterval = TimeSegundos * 1000
    
    SetWindowPos HMen, HWND_TOP, cx, cy, cWidth, cHeight, SWP_NOZORDER

End Function

Function AnchoMonitor() As Long 'Height de la pantalla activa
Dim rec As RECT

    Call GetWindowRect(GetDesktopWindow, rec)
    
    AnchoMonitor = CStr(rec.right - rec.left)

End Function

Function AltoMonitor() As Long 'Widht de la pantalla activa
Dim rec As RECT

    Call GetWindowRect(GetDesktopWindow, rec)
    
    AltoMonitor = CStr(rec.bottom - rec.top)

End Function

Function TamañoBarra() As Long
Dim hwndTrayWnd As Long
Dim res As Long

    hwndTrayWnd& = FindWindow("Shell_TrayWnd", "")
    
    If hwndTrayWnd > 0 Then
        res = GetWindowRect(hwndTrayWnd, buffRECT)
        If res > 0 Then
            TamañoBarra = CStr(buffRECT.bottom - buffRECT.top)
        End If
    End If
 
End Function


