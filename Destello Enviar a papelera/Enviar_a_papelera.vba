'El código está diseñado para trabajar desde un módulo de formulario
Option Compare Database
Option Explicit

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_DELETE = &H3

Private Const FOF_ALLOWUNDO = &H40

Public Sub EnviarAPepelera(ByVal Fichero As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-enviar-fichero-a-la-papelera
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : EnviarAPepelera
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Propósito         : enviar un fichero a la papelera de reciclaje
' Retorno           : cero si tiene éxito y distinto de cero si no lo tiene
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     fichero      Obligatorio     Dirección completa del fichero que queremos eliminar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shfileoperationa
'                     https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shfileopstructa
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'                    Me.pruebatxt.value contiene la dirección completa del fichero que queremos eliminar
'
'Sub EnviarAPepelera_test()
'
'   EnviarAPepelera Me.txtPrueba.Value
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim SHFileOp As SHFILEOPSTRUCT
Dim retorno As Long

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = Fichero
        .fFlags = FOF_ALLOWUNDO
    End With
    
    retorno = SHFileOperation(SHFileOp)
    
End Sub

Private Sub btnKill_Click()

Kill Me.txtPrueba.Value

End Sub

Private Sub btnPapelera_Click()

EnviarAPepelera (Me.txtPrueba.Value)

End Sub

