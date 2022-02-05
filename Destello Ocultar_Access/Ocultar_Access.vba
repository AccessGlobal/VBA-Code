'Declarar a nivel de módulo
Public Declare PtrSafe Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "USER32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Function OcultarAccess(Ocultar As Boolean) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-ocultar-access/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : OcultarAccess
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Actualizado       : marzo 2012
' Propósito         : Ocultar la ventana de Access
' Retorno           : verdadero / falso según se oculte o no
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Ocultar          Obligatorio    El valor Boolean especifica si mostramos o no la ventana
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : SUSTITUIR_POR_EL_NOMBRE_DE_REFERENCIA (C:\ Sustituir_por_la_ruta_completa_de_la_librería)
' Importante        : SUSTITUIR_POR_UNA_BREVE_DESCRIPCIÓN_DE_LA_NOTA_IMPORTANTE
' Más información   : SUSTITUIR_POR_UNA_URL
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub OcultarAccess_test()
'
'   OcultarAccess true
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim lngHWnd As Long
Dim bytNivel As Byte
lngHWnd = Application.hWndAccessApp
bytNivel = IIf(Ocultar, 0, 255)
SetWindowLong lngHWnd, GWL_EXSTYLE, GetWindowLong(lngHWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes lngHWnd, 0, bytNivel, LWA_ALPHA
OcultarAccess = True
End Function