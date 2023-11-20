Option Compare Database
Option Explicit

Function BrowseFolder(Inicio As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-busca-carpeta-de-otra-forma/
'                     Destello formativo 381
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : BrowseFolder
' Autor original    : Alba Salvá | albasalva@access-global.net
' Creado            : desconocido
' Propósito         : Seleccionar una carpeta en el selector de Windows.
' Retorno           : Dirección de la carpeta seleccionada
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte           Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Path_inicial    Obligatorio    Dirección en la que comenzará la búsqueda
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/api/access.application.filedialog
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pégalo en el editor de VBA.
'
'Sub BrowseFolder_test()
'Dim carpeta As String
'Dim MiRuta As String
'
'    MiRuta = "C:\"
'    carpeta = BrowseFolder(MiRuta)
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim BF As Object
Const msoFileDialogFolderPicker = 4
    
'Seleccionamos una carpeta
    Set BF = FileDialog(msoFileDialogFolderPicker)
    
        With BF
            .AllowMultiSelect = False
            .ButtonName = "Seleccionar"
            .InitialFileName = Inicio
            .Title = "Buscar Carpeta"
            If .Show = -1 Then
                 BrowseFolder = .SelectedItems(1)
            End If
        End With
    
    Set BF = Nothing
    
    Debug.Print BrowseFolder
       
End Function
