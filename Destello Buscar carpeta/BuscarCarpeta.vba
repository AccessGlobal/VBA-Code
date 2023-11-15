'Módulo estándar: modCarpetas'
Option Compare Database
Option Explicit

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                                      ByVal lpOperation As String, _
                                                                                      ByVal lpFile As String, _
                                                                                      ByVal lpParameters As String, _
                                                                                      ByVal lpDirectory As String, _
                                                                                      ByVal nShowCmd As Long) As Long

Function Busca_Carpeta(ByVal Titulo As String, ByVal Path_inicial As Variant) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-busca-carpetas/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Busca_Carpeta
' Autor original    : desconocido
' Creado            : desconocido
' Propósito         : Seleccionar una carpeta en el selector de Windows.
' Retorno           : Dirección de la carpeta seleccionada
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte           Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Titulo       Obligatorio       Título del selector
'                  Path_inicial    Obligatorio       Dirección en la que comenzará la búsqueda
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objShell As Object, objFolder As Object, o_Carpeta As Object
    
    Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.BrowseForFolder(0, Titulo, 4000, Path_inicial)
            Set o_Carpeta = objFolder.Self
             
                Busca_Carpeta = o_Carpeta.Path
                o_Carpeta.Modal = True
            
            Set objShell = Nothing
        Set objFolder = Nothing
    Set o_Carpeta = Nothing

End Function