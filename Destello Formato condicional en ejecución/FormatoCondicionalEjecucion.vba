'Módulo estándar de un formulario
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-formato-condicionalen-tiempo-de-ejecucion/
'                     Destello formativo 358
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : btnFormat
' Autor original    : Luis Viadel
' Creado            : 05/10/2023
' Adaptado          : Luis Viadel | https://access-global.net
' Propósito         : Manejar formatos condicionales mediante vba en tiempo de ejecución
'---------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/api/access.formatcondition
' Elementos         : formulario con dos campos de texto (cueact y cueal) con diferentes formatos condicionales
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btnFormat_Click()
Dim cond As Integer

'Mostrar formato condicional
    With Me.cueact
        Debug.Print "Campo: cueact"
        Debug.Print "-------------------"
        For cond = 0 To .FormatConditions.Count - 1
            Debug.Print "Condición: " & cond
            Debug.Print "Color de fondo: " & .FormatConditions(cond).BackColor
            Debug.Print "Color de texto: " & .FormatConditions(cond).ForeColor
        Next
    End With
    Debug.Print ""
    
    With Me.cueal
        Debug.Print "Campo: cueal"
        Debug.Print "-------------------"
        For cond = 0 To .FormatConditions.Count - 1
            Debug.Print "Condición: " & cond
            Debug.Print "Expresión 1: " & .FormatConditions(cond).Expression1
            Debug.Print "Expresión 2: " & .FormatConditions(cond).Expression2
            Debug.Print "Color de fondo: " & .FormatConditions(cond).BackColor
            Debug.Print "Color de texto: " & .FormatConditions(cond).ForeColor
        Next
    End With

'Modificar formato condicional
    With Me.cueact
        For cond = 0 To .FormatConditions.Count - 1
            .FormatConditions(cond).BackColor = vbWhite
            .FormatConditions(cond).ForeColor = vbWhite
        Next
    End With

'Eliminar formato condicional
    Me.cueact.FormatConditions.Delete

'Añadir formato condicional
    Me.cueact.FormatConditions.Add acFieldValue, acEqual, -1
    Me.cueact.FormatConditions(0).BackColor = vbRed
    Me.cueact.FormatConditions(0).ForeColor = vbRed
       
    Me.Refresh
    
End Sub
