Option Compare Database
Option Explicit

Public Declare PtrSafe Function FtpGetFileA Lib "wininet.dll" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare PtrSafe Function FtpPutFileA Lib "wininet.dll" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare PtrSafe Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Long
Public Declare PtrSafe Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

Public Const FTP_TRANSFER_TYPE_UNKNOWN As Long = 0
Public Const INTERNET_FLAG_RELOAD As Long = &H80000000


Public Function FtpUpload(ByVal strLocalFile As String, ByVal strRemoteFile As String, ByVal strHost As String, ByVal lngPort As Long, ByVal strUser As String, ByVal strPass As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-ftp-con-access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FtpUpload
' Autor original    : desconocido
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Fecha             : desconocida
' Propósito         : subir un fichero a un servidor mediate el protocolo ftp
' Retorno           : verdadero/falso según haya tenido éxito o no la transferencia
' Argumento/s       : la sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strLocalFile      Obligatorio    Ruta completa del fichero que queremos subir al servidor
'                     strRemoteFile     Obligatorio    Nombre del fichero de destino
'                     strHost           Obligatorio    URL del servidor
'                     lngPort           Obligatorio    Puerto de comunicaciones utilizado (generalmente el 21)
'                     strUser           Obligatorio    Usuario de acceso al servidor
'                     strPass           Obligatorio    Contraseña de acceso al servidor
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetconnecta
'                     https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetopena
'                     https://https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-internetclosehandle
'                     https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-ftpgetfilea
'                     https://learn.microsoft.com/en-us/windows/win32/api/wininet/nf-wininet-ftpputfilea
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, rellena los datos de tu servidor
'                     y pulsa F5 para ver su funcionamiento.
'
'Sub FtpUpload_test()
'
'   Call FtpUpload(ruta de origen, nombreFichero en destino, "URL servidor", "puerto", "usuarioftp", "pass")
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim hOpen As Long
Dim hConn As Long

    hOpen = InternetOpen("FTPGET", 1, vbNullString, vbNullString, 1)
    hConn = InternetConnect(hOpen, strHost, lngPort, strUser, strPass, 1, 0, 2)
    
    If FtpPutFileA(hConn, strLocalFile, strRemoteFile, FTP_TRANSFER_TYPE_UNKNOWN Or INTERNET_FLAG_RELOAD, 0) Then
        FtpUpload = True
    Else
        FtpUpload = False
    End If
    
'Close connections
    InternetCloseHandle hConn
    InternetCloseHandle hOpen

End Function

Public Function FtpDownload(ByVal strRemoteFile As String, ByVal strLocalFile As String, ByVal strHost As String, ByVal lngPort As Long, ByVal strUser As String, ByVal strPass As String) As Boolean
Dim hOpen   As Long
Dim hConn   As Long

    hOpen = InternetOpen("FTPGET", 1, vbNullString, vbNullString, 1)
    hConn = InternetConnect(hOpen, strHost, lngPort, strUser, strPass, 1, 0, 2)
    
    If FtpGetFileA(hConn, strRemoteFile, strLocalFile, 1, 0, FTP_TRANSFER_TYPE_UNKNOWN Or INTERNET_FLAG_RELOAD, 0) Then
        FtpDownload = True
    Else
        FtpDownload = False
    End If

'Close connections
    InternetCloseHandle hConn
    InternetCloseHandle hOpen
    
End Function
