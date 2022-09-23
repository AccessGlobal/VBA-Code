Option Compare Database
Option Explicit

Private Declare Function AddFontResource Lib "gdi32.dll" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Sub NuevaFuente()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-nueva-fuente
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : NuevaFuente
' Autor original    : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : noviembre 2019
' Propósito         : instalar una nueva fuente en el sistema
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/es-es/windows/win32/api/wingdi/nf-wingdi-addfontresourcea?redirectedfrom=MSDN
'                   : https://access-global.net/vba-metodo-fileexists/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, seleccionar una fuente que no se encuentre
'                     en el sistema y pulsar F5 para ver su funcionamiento.
'
' Sub NuevaFuente_test()
'
'       AddFontResource(Application.CurrentProject.Path & "\Fonts\MiFuente.ttf")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ruta As String
Dim fso, Archivo
Dim result As Long

ruta = "C:\Windows\Fonts\ean13.ttf"

'Comprueba si ya existe el fichero mediante el método FileExists del objeto FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(ruta) Then
            Exit Sub
        Else
            result = AddFontResource(ruta)
        End If
    Set fso = Nothing
    
End Sub

