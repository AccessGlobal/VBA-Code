Option Compare Database
Option Explicit

Public Sub ImgageView(ImgPath As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-utilizar-una referencia-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ImgageView
' Autor original    : zxs23 (https://www.forosdelweb.com/f69/como-ejecutar-imagen-con-visor-imagenes-windows-935362/)
' Adaptado          : Luis Viadel
' Creado            : enero 2023
' Propósito         : mostrar una imagen en el visor de imágenes de Windows
' Argumento/s       : La sintaxis de la rutina consta del siguiente argumento:
'                     Nombre          Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     ImgPath         Obligatorio      Dirección completa del fichero de imagen que queremos visualizar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-loadlibraryw
'                   : https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-getprocaddress
'                   : https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-callwindowproca
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub ImgageView_test()
'Dim ImgPath As String
'
'    ImgPath = "Dirección de la imagen"
'
'    Call ImgageView(ImgPath)
'
'End Sub
'
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Lib As Long
Dim LibAdd As Long
    
    Lib = LoadLibrary("shimgvw")

    LibAdd = GetProcAddress(Lib, "imageview_fullscreenW")

    CallWindowProc LibAdd, 0&, 0&, StrPtr(ImgPath), 0&
    
'Liberar la librería
    FreeLibrary Lib
    
End Sub

