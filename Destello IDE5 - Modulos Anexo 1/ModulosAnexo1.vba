
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-otros-ejemplos-tipos-de-modulo-segun-contenido/
'                     Destello formativo 272
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AnexoModulos
' Autor original    : Alba Salvá
' Creado            : 22/02/2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : Se trata de un par de funciones para inverstigar el contenido de los módulos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------



'Colocar el siguiente código en el load de un formulario que contiene un listbox (lstMods)
Private Sub Form_Load()
Dim arrMods As Variant
Dim varMod As Variant
Dim objMod As VBIDE.VBComponent
    
    Me.lstMods.RowSource = ""
    Me.lstMods.Requery
    
    arrMods = Array("Modulo_Vacio", "Modulo_Cabecera", "Modulo_Procedimiento")
    
    For Each varMod In arrMods
        Set objMod = Application.VBE.ActiveVBProject.VBComponents(CStr(varMod))
        If ModuloVacio(objMod) Then
            Me.lstMods.AddItem "El módulo '" & CStr(varMod) & "' está vacío"
        Else
            If TieneProcedimientos(objMod) Then
                Me.lstMods.AddItem "El módulo '" & CStr(varMod) & "' tiene procedimientos"
            Else
                Me.lstMods.AddItem "El módulo '" & CStr(varMod) & "' sólo tiene cabecera"
            End If
        End If
    Next
    
    Set objMod = Nothing

End Sub

'Copia y pega este código en un módulo estándar
Option Compare Database
Option Explicit

Public Function ModuloVacio(vbc As VBIDE.VBComponent) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-otros-ejemplos-tipos-de-modulo-segun-contenido/
'                     Destello formativo 272
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ModuloVacio
' Autor original    : Alba Salvá
' Creado            : 22/02/2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : comprueba si el módulo está vacío, sólo se considera 
'                    vacío si no tiene líneas o se trata de las declaraciones
'                    Option y líneas comentadas
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim lngLine As Long
Dim strLine As String
 
    ModuloVacio = False
 
    If vbc.CodeModule.CountOfLines = vbc.CodeModule.CountOfDeclarationLines Then ' Compruebo que tiene líneas
        For lngLine = 1 To vbc.CodeModule.CountOfLines
            strLine = Trim(vbc.CodeModule.Lines(lngLine, 1))
            If Len(strLine) > 0 Then                        ' Verifico si la línea tiene caracteres
                If Left(strLine, 6) <> "Option" Then        ' Descarto las líneas de "Option..."
                    If Left(strLine, 1) <> "'" Then         ' Descarto las líneas de comentario con comilla simple
                        If Left(strLine, 3) <> "Rem" Then   ' Descarto las líneas de comentario con "REM"
                            Exit Function                   ' Salgo con falso
                        End If
                    End If
                End If
            End If
        Next
        ModuloVacio = True                             'Indico que está vacío
    End If
 
End Function

Public Function TieneProcedimientos(vbc As VBIDE.VBComponent) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-otros-ejemplos-tipos-de-modulo-segun-contenido/
'                     Destello formativo 272
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : TieneProcedimientos
' Autor original    : Alba Salvá
' Creado            : 22/02/2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : comprueba si el módulo está tiene algún procedimiento
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------

    With vbc.CodeModule
        TieneProcedimientos = Not (.CountOfDeclarationLines = .CountOfLines)
    End With
 
End Function

