'Escribe en un módulo cualquiera
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal wFlags As Long) As Long
Public Declare PtrSafe Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fwinIni As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const WS_EX_LAYERED = &H80000
Public Function Aplicar_Transparencia(ByVal hWnd As Long, valor As Byte) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/aplicar-transparencia-a-un-formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Aplicar transparencia
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Actualizado       : febrero 2015
' Propósito         : Aplicar un grado de transparencia a un form con valores entre 1 y 255 utilizando la API de Windows
' Retorno           : Valor long de la transparencia
'                     0 equivale a que se aplica transparencia según se ha seleccionado
'                     1 equivale a sin transparencia por valores fuera de rango
'                     2 se ha producido un error y no se aplica transparencia
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     hWnd              Obligatorio        Identificador de windows del objeto formulario
'                     valor             Obligatorio        Grado de transparencia (entre 1 y 255)
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en el evento de carga de cualquier formulario que desees.
'
'                     Private Sub Form_load()
'
'                         Call Aplicar_Transparencia(Me.hWnd, 120)
'
'                      End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Msg As Long
On Error Resume Next
If valor < 0 Or valor > 255 Then
Aplicar_Transparencia = 1
Else
Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
Msg = Msg Or WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hWnd, 0, valor, LWA_ALPHA
Aplicar_Transparencia = 0
End If
If Err Then
Aplicar_Transparencia = 2
End If
End Function