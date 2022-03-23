Option Compare Database
Option Explicit

Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Global Const SND_ASYNC = &H1
Global Const SND_FILENAME = &H20000
Global Const SND_NODEFAULT = &H2

Public Function Sonido(file As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-sonido-en-mi-app
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Sonido
' Autor original    : Luis Viadel
' Creado            : octubre 2009
' Propósito         : emitir un sonido en ciertos eventos de la aplicación
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     file          Obligatorio      Nombre del fichero de sonido que queremos reproducir
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/previous-versions//dd743680(v=vs.85)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub test_sonido()
'
'Call Sonido("MiFicheroDeSonido")
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------------------------------------------
Call PlaySound(CurrentProject.Path & "\Sonido\" & file & ".wav", 0&, SND_FILENAME Or SND_NODEFAULT)

End Function