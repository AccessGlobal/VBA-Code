'Módulo estándar 
Public Function IsBlank(arg As Variant) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-esta-vacio-es-blanco-es-nulo-se-ha-perdido/
'                     Destello formativo 361
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : IsBlank
' Autor original    : Renaud
' Fuente original   : http://blog.nkadesign.com/2009/09/09/access-checking-blank-variables/
' Creado            : 2009
' Propósito         : chequear cadenas y objetos para detectar valores blancos y null
' Mas información   : https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/vartype-function
'                     https://support.microsoft.com/en-us/office/ismissing-function-22286f0f-d1e7-4ce4-96d0-7691a3944bf1
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Sub test()
' Dim blank as boolean
'
'   blank=IsBlank(Mi Objeto)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    
    Select Case VarType(arg)
        Case vbEmpty
            IsBlank = True
        Case vbNull
            IsBlank = True
        Case vbString
            IsBlank = (arg = VBA.vbNullString Or arg = vbNullChar)
        Case vbObject
            IsBlank = (arg Is Nothing)
        Case Else
            IsBlank = IsMissing(arg)
    End Select

End Function


'Ejemplo de uso en matrices
Public Function ArrayBlank(ByRef MyArray) As Boolean
    
    On Error Resume Next
    
    ArrayBlank = Not IsBlank(MyArray(0))
    
End Function

Sub MatrizBlank_test()

'....disponemos de una matriz de ficheros a la que hemos llamado MyArray y queremos
' comporbar que no hay ningún elemento vacío

    If ArrayBlank(Files) Then
        For i = 0 To UBound(Files)
            ruta = Files(i)
        Next
    End If

End Sub

'Ejemplo de uso en cuadros de texto
Sub textBox_test()
    
    If Not (IsBlank(MiTextBox)) Then
        miVarible = MiTextBox.value
    Else
    '...
    End If

End Sub

'Ejemplo de uso en variables
Sub variables_test()

    If IsBlank(MiVariable) Then
    '...
    Else
    '...
    End If
       
End Sub
