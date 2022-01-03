Public Function mcblnMkDir(Optional ByVal strPathName As String, Optional ByVal blnRestore As Boolean) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/eliminar-una-carpeta-y-todo-su-contenido/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcblnMkDir
' Autor original    : Luis Viadel
' Autor             : Rafael Andrada .:McPegasus:. | BeeSoftware | rafael.andrada@access-global.net
' Actualizado       : 15/07/2019 12:54
' Propósito         : Restaura la carpeta Temp cuando termina el proceso.
' Retorno           : False en caso de haberse producido cualquier error no esperado.
' Argumento/s       : La sintaxis del Procedimiento o Función consta de/los siguiente/s argumento/s:
'                     Parte             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strPathName       Opcional    Nombre de la carpeta a eliminar.
'                     blnRestore        Opcional    En caso de querer crear la carpeta tras su eliminación y por tanto de su contenido.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : Microsoft Scripting Runtime (c:\Windows\SysWOW64\scrrun.dll)/>
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Copia el bloque siguiente al portapapeles y pega en el editro de VBA para ver un ejemplo de funcionamiento.
'Sub mcblnMkDir_test()
'   Debug.Print mcblnMkDir ("C:\temp")
'   Debug.Print mcblnMkDir(, True)
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
'    Dim fso                                         As FileSystemObject
'    Dim fsoFolder                                   As Folder
Dim fso                                         As Object
Dim fsoFolder                                   As Object
On Error GoTo mcblnMkDir_ErrorHandler                                  'Controlador de errores por On Error Goto. Mc 13/07/2021, 14/06/2021.
If strPathName = "" Then
If Right(strPathName, 1) = "\" Then
'No debe de tener la barra final, en ese caso se produce un error de que no se ha encontrado la ruta de acceso.
strPathName = Left(strPathName, Len(strPathName) - 1)
End If
Set fso = CreateObject("Scripting.FileSystemObject")                'https://vba846.wordpress.com/file-system-object-para-vba/
fso.DeleteFolder strPathName, True
If blnRestore Then
Set fsoFolder = fso.CreateFolder(strPathName)
End If
End If
mcblnMkDir_Exit:
Set fso = Nothing
Set fsoFolder = Nothing
On Error GoTo 0
Exit Function
mcblnMkDir_ErrorHandler:
Select Case Err.Number
Case 76                                     'No se ha encontrado la ruta de acceso.
Resume Next                             'Para volver a la línea siguiente donde se produce el error.
Case Else                                   'Cazar todos aquellos errores inesperados.
'La línea que se muestra en la ventana de Inmediato con el Debug.Print, está preparada por su tabulación para sustituir por un Case 0.
Debug.Print "        Case " & Err.Number & "                                     '" & Err.Description & "."
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mcblnMkDir.", vbCritical
mcblnMkDir = True
'Stop
'Resume                                              'Para volver al punto donde se produce el error.
'Resume Next                                         'Para volver a la línea siguiente donde se produce el error.
End Select
Resume mcblnMkDir_Exit             'Ir al procedimiento de salida.
End Function
