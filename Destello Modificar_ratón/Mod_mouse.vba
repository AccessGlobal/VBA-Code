Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-crea-tu-propio-icono-de-raton
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : modMousePointer (change mouse icon)
' Autor original    : Terry Kreft
' Adaptado por      : Luis Viadel
' Actualizado       : mayo 2015
' Propósito         : Cambiar el icono del ratón en tiempo de ejecución, por cualquier imagen que el usuario cree
' ¿Cómo funciona?   : hay dos funciones
'                     SetMouseCursorFromFile se utiliza junto con una ruta a un archivo .ico
'                     SetMouseCursor se utiliza junto con una de las siguientes constantes:
' Más información   : http://www.mvps.org/access/api/api0044.htm
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public Const IDC_APPSTARTING As Long = 32650&
Public Const IDC_HAND As Long = 32649&
Public Const IDC_ARROW As Long = 32512&
Public Const IDC_CROSS As Long = 32515&
Public Const IDC_IBEAM As Long = 32513&
Public Const IDC_ICON As Long = 32641&
Public Const IDC_NO As Long = 32648&
Public Const IDC_SIZE As Long = 32640&
Public Const IDC_SIZEALL As Long = 32646&
Public Const IDC_SIZENESW As Long = 32643&
Public Const IDC_SIZENS As Long = 32645&
Public Const IDC_SIZENWSE As Long = 32642&
Public Const IDC_SIZEWE As Long = 32644&
Public Const IDC_UPARROW As Long = 32516&
Public Const IDC_WAIT As Long = 32514&

Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare PtrSafe Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Function SetMouseCursor(CursorType As Long)
Dim lngRet As Long

    lngRet = LoadCursorBynum(0&, CursorType)
    lngRet = SetCursor(lngRet)

End Function

Public Function SetMouseCursorFromFile(strPathToCursor As String)
Dim lngRet As Long

    lngRet = LoadCursorFromFile(strPathToCursor)
    lngRet = SetCursor(lngRet)

End Function

Sub test_cursor()
Dim strIconPath As String

strIconPath = Application.CurrentProject.Path & "\Miicono.ico"

If Len(Dir$(strIconPath)) > 0 Then
    SetMouseCursorFromFile strIconPath
Else
    SetMouseCursor IDC_CROSS
End If

End Sub
