'Módulo de formulario
Option Compare Database
Option Explicit
Private Sub btnGetAttr_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-getattr
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : btnGetAttr
' Autor original    : Luis Viadel
' Adaptado por      : mayo 2021
' Propósito         : obtener los atributos de un fichero en tiempo de ejecución
' Retorno           : valor integer que indica los atributos del fichero
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getattr-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub EstaAbierto2_test()
' Dim resultado As Integer
'
'     On Error GoTo LinSalir
'     Me.txtResultado = "El fichero existe
'     Exit Sub
' LinSalir:
'     MsgBox "El fichero no existe"
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim resultado As Integer
Dim txtResultado As String
    On Error GoTo LinSalir
        resultado = GetAttr(Me.txtPrueba)
        Select Case resultado
            Case 0
                txtResultado = "Normal"
            Case 1
                txtResultado = "Sólo lectura"
            Case 2
                txtResultado = "Oculto"
            Case 4
                txtResultado = "Fichero de sistema" '(no disponible en Mac)
            Case 16
                txtResultado = "Directorio o carpeta"
            Case 32
                txtResultado = "El fichero ha cambiado desde el último backup" ' (no disponible en Mac)
            Case 64
                txtResultado = "El fichero especificado es un alias" '(sólo disponible en mac)
        End Select
        Me.txtResultado = "El fichero existe. " &  txtResultado
    Exit Sub
LinSalir:
    MsgBox "El fichero no existe"
End Sub