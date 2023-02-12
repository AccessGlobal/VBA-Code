Option Compare Database
Option Explicit

Public Sub ListadoProcedimientos(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-listar-procedimientos/
'                     Destello formativo 264
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoProcedimientos
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : mostrar un listado con todos los procedimientos de nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/language/glossary/vbe-glossary#code-pane
'                     https://learn.microsoft.com/en-us/office/vba/api/access.module
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ListadoProcedimientos Me
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent
Dim lngStartLine As Long
Dim lngProcTyp As Long
    
    frm.lstProc.RowSource = ""
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        frm.lstProc.AddItem vbc.Name
        
        With vbc.CodeModule
'Propiedades de CodeModule
'            Debug.Print "Módulo: " & .CodePane.CodeModule
'            Debug.Print "Nº de líneas en las declaraciones: " & .CountOfDeclarationLines
'            Debug.Print "Nº de líneas en el módulo: " & .CountOfLines
'            Debug.Print "Contenido de la línea nº 21: " & .Lines(21, 1)
'            Debug.Print "Objeto principal: " & .Parent.Name
'            Debug.Print "El cuerpo del procedimiento '" & .ProcOfLine(13, vbext_pk_Proc) & "' comienza en la línea " & .ProcBodyLine("lstProc_DblClick", vbext_pk_Proc)
'            Debug.Print "El procedimiento '" & .ProcOfLine(13, vbext_pk_Proc) & "' contiene "; .ProcCountLines("lstProc_DblClick", vbext_pk_Proc) & " líneas"
'            Debug.Print "El procedimiento que contiene la línea 6 es " & .ProcOfLine(6, vbext_pk_Proc)
'            Debug.Print "El procedimiento lstProc_DblClick comienza en la línea " & .ProcStartLine("lstProc_DblClick", vbext_pk_Proc)
'
            lngStartLine = .CountOfDeclarationLines + 1
          
            Do Until lngStartLine >= .CountOfLines
'Obtenemos el nombre del procedimiento con ProcOfLine a través del número de línea
                frm.lstProc.AddItem String(8, "-") & " " & .ProcOfLine(lngStartLine, lngProcTyp) & _
                                    " (" & Choose(lngProcTyp + 1, "Procedimiento / Evento", "Let", "Set", "Get") & ")"
                lngStartLine = lngStartLine + .ProcCountLines(.ProcOfLine(lngStartLine, lngProcTyp), lngProcTyp)
            Loop
        
        End With
        
    Next
       
End Sub

Public Sub ListadoProcedimientosForm()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-listar-procedimientos/
'                     Destello formativo 264
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoProcedimientosForm
' Creado            : Luis Viadel
' Propósito         : mostrar un listado todos los procedimientos de nuestro programa en un UserForm de Acccess
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private UserForm_Initialize()
'
'    ListadoProcedimientosForm
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent
Dim lngStartLine As Long
Dim lngProcTyp As Long

    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        If vbc.Name <> "procUserForm" Then
            procUserForm.lstProc.AddItem vbc.Name
            
            With vbc.CodeModule
        
              lngStartLine = .CountOfDeclarationLines + 1
            
              Do Until lngStartLine >= .CountOfLines
                  procUserForm.lstProc.AddItem String(8, "-") & " " & .ProcOfLine(lngStartLine, lngProcTyp) & _
                                      " (" & Choose(lngProcTyp + 1, "Procedimiento / Evento", "Let", "Set", "Get") & ")"
                  lngStartLine = lngStartLine + .ProcCountLines(.ProcOfLine(lngStartLine, lngProcTyp), lngProcTyp)
              Loop
        
            End With
        
        End If
    
    Next
        
End Sub
