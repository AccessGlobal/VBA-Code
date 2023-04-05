Public Function TotalLineas(vbproject As String, vbMod As String) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-esta-linea-es-un-comentario/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : TotalLineas
' Autor original    : Alba Salvá
' Creado            : 22/01/2022
' Adaptado por      : Luis Viadel
' Propósito         : Obtener las líneas de código real sin contar las líneas comentadas
'					Destello 301
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la parte superior de este módulo.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Sub TotalLineas_test()
'
'    Debug.Print TotalLineas("MiProyecto", "MiModulo")
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent
Dim vbcProjects As VBIDE.VBProjects
Dim vbcProject As VBIDE.vbproject
Dim lngStartLine As Long, lngLastLine As Long
Dim i As Long
Dim lngComment As Long
   
    Set vbcProjects = Application.VBE.VBProjects
        For Each vbcProject In vbcProjects
            If vbcProject.Name = vbproject Then
                For Each vbc In vbcProject.VBComponents
                    If vbc.Name = vbMod Then
                        With vbc.CodeModule
                            lngLastLine = .CountOfLines
                            lngStartLine = .CountOfDeclarationLines + 1
                            For i = lngStartLine To lngLastLine
'Comprobamos si lla línea es un comentario
                                If vbc.CodeModule.Find("'", i, 1, i, 10) = True Then lngComment = lngComment + 1
                                If vbc.CodeModule.Find("Rem", i, 1, i, 10) = True Then lngComment = lngComment + 1
                            Next
                        End With
                    End If
                Next
            End If
        Next
    Set vbcProjects = Nothing
 
    TotalLineas = lngLastLine - lngComment
 
End Function
