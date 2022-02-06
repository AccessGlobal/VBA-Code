''Declarar a nivel de módulo
Public Declare PtrSafe Function apiGetSystemMenu Lib "USER32" Alias "GetSystemMenu" (ByVal hwnd As Long, ByVal flag As Long) As Long
Public Declare PtrSafe Function apiEnableMenuItem Lib "USER32" Alias "EnableMenuItem" (ByVal hMenu As Long, ByVal wIDEnableMenuItem As Long, ByVal wEnable As Long) As Long
Const MF_BYCOMMAND = &H0&
Const MF_DISABLED = &H2&
Const MF_ENABLED = &H0&
Const MF_GRAYED = &H1&
Const SC_CLOSE = &HF060&
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOMOVE = &H2
Const SWP_FRAMECHANGED = &H20
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_SYSMENU = &H80000
Function ActivarBotonCerrar(bEnable As Boolean, Optional ByVal lhWndTarget As Long = 0) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-bloquear-cerrar-access/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título original   : EnableDisableCloseButton
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Actualizado       : abril 2010
' Propósito         : Impedir el cierre de la aplicación de forma accidental
' Retorno           : valor long de la acción
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     bEnable           Obligatorio        El valor Boolean especifica si queremos bloquear el botón (false) o activarlo (true)
'                     lhWndTarget       Opcional           Hace referencia a la ventana de Access en la que queremos actuar
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub ActivarBotonCerrar_test()
'
'    ActivarBotonCerrar False
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim lhWndMenu As Long
Dim lReturnVal As Long
Dim lAction As Long
Const wFlags = SWP_NOSIZE + SWP_NOZORDER + SWP_FRAMECHANGED + SWP_NOMOVE
Const FLAGS_COMBI = WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_SYSMENU
lhWndMenu = apiGetSystemMenu(IIf(lhWndTarget = 0, Application.hWndAccessApp, lhWndTarget), False)
If lhWndMenu <> 0 Then
If bEnable Then
lAction = MF_BYCOMMAND Or MF_ENABLED
Else
lAction = MF_BYCOMMAND Or MF_DISABLED Or MF_GRAYED
End If
lReturnVal = apiEnableMenuItem(lhWndMenu, SC_CLOSE, lAction)
End If
ActivarBotonCerrar = lReturnVal
End Function
