Public Function mcblnCreateFolders(ByVal strFullPathWhitDriver As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/crear-carpetas-encadenadas/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcblnCreateFolders
' Autor             : Rafael Andrada .:McPegasus:. | BeeSoftware | rafael.andrada@access-global.net
' Actualizado       : 10/10/2021
' Propósito         : Crear varias carpetas que dependan de una o de una unidad de disco. No probado con unidad de red.
' Retorno           : Verdadero si se han creado todas las carpetas con éxito.
' Argumento/s       : La sintaxis del procedimiento o función consta de/los siguiente/s argumento/s:
'                     Parte                             Modo                Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strFullPathWhitDriver       Obligatorio     El valor String especifica la sucesión de carpetas a crear. Ejemplo: C:\Bee-Software\FolderA\FolderB\FolderC\
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim bytCount                                    As Byte
Dim strDrive                                    As String
Dim strPath                                     As String
Dim strPathWork()                               As String
strPath = "C:\Bee-Software\FolderA\FolderB\FolderC\"
strDrive = Left(strPath, 2)
strPathWork = Split(strPath, "\")
strPath = strDrive
bytCount = 1
Do While Not bytCount = UBound(strPathWork) + 1
strPath = strPath & "\" & strPathWork(bytCount)
If Dir(strPath, vbDirectory) = "" Then
MkDir strPath
End If
bytCount = bytCount + 1
Loop
End Function
