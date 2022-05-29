'Código a incluir en el botón de acción que lance la función
Private Sub btnrecupera_Click()
Dim strDireccion As String
Dim strFichero As String
Dim strRecycle As String
Dim strRutaOrigen As String

'Extraemos el nombre de fichero para poder recuperarlo
    strDireccion = Me.txtPrueba.Value
    strFichero = Right(strDireccion, Len(strDireccion) - InStrRev(strDireccion, "\"))

'Localizamos la dirección del fichero en la papelera de reciclaje
    strRecycle = RecuperaPapelera(strFichero)

'Copiamos el fichero en la ruta original
    FileCopy strRecycle, Me.txtPrueba.Value
    Kill strRecycle
    
End Sub

'Código de módulo estándar
Option Compare Database
Option Explicit

Public Function RecuperaPapelera(ByVal fichero As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente             : https://access-global.net/vba-recuperar-el-fichero-enviado-a-la-papelera
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título             : RecuperaPapelera
' Autor original     : Luis Viadel
' Fecha              : marzo 2019
' Propósito          : recuperar un fichero conocido que ha sido enviado a la papelera de reciclaje
' Retorno            : devuelve la dirección completa del fichero en la papelera de reciclaje
' Argumento/s        : La sintaxis del procedimiento o función consta del siguiente argumento:
'                      Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                      fichero       Obligatorio       Nombre del fichero que queremos recuperar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Mas información    : https://docs.microsoft.com/en-us/windows/win32/shell/shell-namespace
'                      https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfile-method
'                      https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test                : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                      portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub recupera_test()
'Dim strRecycle as string
'
'   strRecycle = RecuperaPapelera(Nombre de fichero)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim sh As Object, folder As Object
Dim item As Object

Const BITBUCKET = &HA&

    On Error GoTo LinErr
    
    Set sh = CreateObject("Shell.Application")
    Set folder = sh.Namespace(BITBUCKET)
    
    For Each item In folder.Items

        If InStr(item.Name, fichero) Then
            RecuperaPapelera = item.Path
        End If
    Next
    
    Exit Function
LinErr:
    RecuperaPapelera = ""
    
End Function
