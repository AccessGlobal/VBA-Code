'Módulo estándar : modDSN
Option Compare Database
Option Explicit

Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As Long, ByVal frequest As Long, ByVal lpszdriver As String, ByVal lpszattributes As String) As Long
Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const HKEY_CURRENT_USER = &H80000001
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&

Public Function BuscarDSN(nombreDSN As String) As Boolean
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-mas-cosas-sobre-los-DSN/
'                     Destello formativo 392
'--------------------------------------------------------------------------------------------------------
' Título            : BuscarDSN
' Autor             : Desconocido
' Creado            : Desconocido
' Propósito         : Comprobar la existencia de un DSN
'--------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexa
'                     https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumvaluea
'                     https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
'--------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA.
'--------------------------------------------------------------------------------------------------------
'Sub ListadoDSN_Test()
'Dim nombreDSN as string
'
'    nombreDSN="Mi_DSN"
'
'    If ListadoDSN(nombreDSN) = False Then
'       Aquí escribo mi código
'    End If
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim lngKeyHandle As Long
Dim lngResult As Long
Dim lngCurIdx As Long
Dim strValue As String
Dim lngValueLen As Long
Dim lngData As Long
Dim lngDataLen As Long
Dim strResult As String

    lngResult = RegOpenKeyEx(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", 0&, KEY_READ, lngKeyHandle)
    
    If lngResult <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    lngCurIdx = 0
    Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000

        lngResult = RegEnumValue(lngKeyHandle, lngCurIdx, ByVal strValue, lngValueLen, 0&, REG_DWORD, ByVal lngData, lngDataLen)
        lngCurIdx = lngCurIdx + 1

        If lngResult = ERROR_SUCCESS Then
            strResult = strResult & lngCurIdx & ": " & Left(strValue, lngValueLen) & vbCrLf
            strResult = Left(strValue, lngValueLen) & vbCrLf
            If strResult Like nombreDSN & "*" Then
                BuscarDSN = True
                Call RegCloseKey(lngKeyHandle)
                Exit Function
            End If
        End If
        
     Loop While lngResult = ERROR_SUCCESS
     
     Call RegCloseKey(lngKeyHandle)
      
End Function

Public Sub ListarDSN()
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-mas-cosas-sobre-los-DSN/
'                     Destello formativo 392
'--------------------------------------------------------------------------------------------------------
' Título            : ListarDSN
' Autor             : Desconocido
' Creado            : Desconocido
' Propósito         : Obtener un listado de todos los DSN instalados en el PC
'--------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexa
'                     https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumvaluea
'                     https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
'--------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA.
'--------------------------------------------------------------------------------------------------------
'Sub ListarDSN_Test()
'
'    Call ListarDSN
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim lngKeyHandle As Long
Dim lngResult As Long
Dim lngCurIdx As Long
Dim strValue As String
Dim lngValueLen As Long
Dim lngData As Long
Dim lngDataLen As Long
Dim strResult As String

    lngResult = RegOpenKeyEx(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", 0&, KEY_READ, lngKeyHandle)
    
    lngCurIdx = 0
    Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000

        lngResult = RegEnumValue(lngKeyHandle, lngCurIdx, ByVal strValue, lngValueLen, 0&, REG_DWORD, ByVal lngData, lngDataLen)
        lngCurIdx = lngCurIdx + 1

        strResult = strResult & lngCurIdx & ": " & Left(strValue, lngValueLen) & vbCrLf
        strResult = Left(strValue, lngValueLen) & vbCrLf
        Debug.Print strResult
     
     Loop While lngResult = ERROR_SUCCESS
     
     Call RegCloseKey(lngKeyHandle)
      
End Sub