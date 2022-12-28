Option Compare Database
Option Explicit

Public Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (Optional ByVal pCaller As Long, Optional ByVal szURL As String, Optional ByVal szFileName As String, Optional ByVal dwReserved As Long, Optional ByVal lpfnCB As Long) As Boolean

Public Const urlsite = "https://Mi-URL-de-descarga/"
Public Const nomfic = "Imagen_test.jpg"

Sub DescargaVersion_test()

   Call DescargaVersion(urlsite, nomfic)

End Sub

Public Function DescargaVersion(urlsite As String, nomfic As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-mas-sobre-la-descarga-de-archivos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : DescargaVersion
' Autor             : Luis Viadel | https://cowtechnologies.net
' Fecha             : junio 2015
' Propósito         : descargar un fichero desde una ubicación en Internet
' Retorno           : verdadero/falso según haya tenido éxito o no la descarga
' Argumento/s       : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     urlsite      Obligatorio      url donde se encuentra el fichero
'                     nomfic       Obligatorio      Nombre del fichero de destino
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/ms775123(v=vs.85)
'                     https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-deleteurlcacheentry
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, rellena los datos de la url y los del
'                     fichero que deseas descargar y pulsa F5 para ver su funcionamiento.
'
'Sub DescargaVersion_test()
'
'   Call DescargaVersion(urlsite, nomfic)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ruta As String, URLNueva As String
Dim resultado As Long

    ruta = Application.CurrentProject.Path & "\pruebas\"

    URLNueva = urlsite & nomfic

'Limpia la caché para que pueda descargar el nuevo fichero
      
    DeleteUrlCacheEntry URLNueva
    
    resultado = URLDownloadToFile(0, URLNueva, ruta & nomfic, 0, 0)
   
    If resultado <> 0 Then GoTo LinErr
    
    Call ShellExecute(0&, "open", ruta, 0&, vbNullString, 1&)
        
    DescargaVersion = True
    
    Exit Function

LinErr:
    DescargaVersion = False
    
    Select Case resultado
        Case 5
            MsgBox "No tienes conexión a Internet"
        Case 6
            MsgBox "Se ha producido un error en la descarga. No he podido encontrar el fichero"
        Case Else
            MsgBox "Se ha producido un error en la descarga. No he podido encontrar la carpeta de descarga "
    End Select

End Function
