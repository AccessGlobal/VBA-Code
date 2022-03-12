Public Function EstaAbierto(filename As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-freefile-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : EstaAbierto
' Autor original    : Desconocido
' Adaptado por      : Luis Viadel
' Propósito         : Conocer el estado de un fichero en tiempo de ejecución
' Retorno           : Valor booelano indicando el estado (abierto = True, cerrado = False)
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     filename      Obligatorio    Dirección del fichero que queremos evaluar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/freefile-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub EstaAbierto_test()
'Dim strdoc as string
'
'strdoc = "C:/Documentos/midocumento.docx"
'    If EstaAbierto (strdoc) then
'       debug.print "Está abierto"
'    Else
'       debug.print "Está cerrado"
'    End if
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim filenum As Integer, ErrNum As Integer

On Error Resume Next

filenum = FreeFile()

'Intenta abrir el fichero para escritura y lo bloquea
Open filename For Input Lock Read As #filenum
'Cierra el fichero y captura el error
    Close filenum
    ErrNum = Err
    On Error GoTo 0
    Select Case ErrNum
        Case 0
            EstaAbierto = False
        Case 70
            EstaAbierto = True
        Case Else
            Error ErrNum
    End Select
    
End Function
