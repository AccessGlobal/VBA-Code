Option Compare Database
Option Explicit

Enum CompType
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_Document = 100
End Enum

#If False Then
    Dim vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document
#End If

Public Sub ListadoModulos(frm As Form, Optional Filtro As CompType)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-listar-modulos/
'                     Destello formativo 259
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoModulos
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : mostrar un listado todos los módulos de nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/collections-visual-basic-add-in-model
'                     https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ListadoModulosListadoModulos Me,NombreFiltro
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent
Dim strType As Variant
     
    strType = Array("", "Módulo estándar", "Módulo de clase", "MS Form")
    
    ReDim Preserve strType(100)
    
    strType(100) = "Formulario | Informe"
    
    frm.lstMods.RowSource = ""
    frm.lstMods.Requery
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        If vbc.Name <> "modListado" Then
'Cada una de las propiedades
'            Debug.Print "Nombre: " & vbc.Name
'            Debug.Print "Código: " & vbc.CodeModule
'            Debug.Print "Id del diseñador: " & vbc.DesignerID
'            Debug.Print "¿Diseñador abierto?: " & vbc.HasOpenDesigner
'            Debug.Print "¿Se ha guardado?: " & vbc.Saved
'            Debug.Print "Tipo: " & strType(vbc.Type)
'            Debug.Print vbNewLine
            
'Los métodos
'            Debug.Print vbc.Activate
'            Debug.Print vbc.DesignerWindow
'            Debug.Print vbc.Export
'            Debug.Print vbc.Designer

            If Filtro > 0 And vbc.Type = Filtro Then
                frm.lstMods.AddItem vbc.Name
            ElseIf Filtro = 0 Then
                frm.lstMods.AddItem vbc.Name & "  |  " & strType(vbc.Type)
            End If
        End If
    
    Next
    
    Set vbc = Nothing
    
End Sub

Public Sub ListadoModulosForm()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-listar-modulos/
'                     Destello formativo 259
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoModulosForm
' Creado            : Luis Viadel
' Propósito         : mostrar un listado todos los módulos de nuestro programa en un UserForm de Acccess
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private UserForm_Initialize()
'
'    ListadoModulosForm
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent
    
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        If vbc.Name <> "modUserForm" Then
            modUserForm.lstModulos.AddItem vbc.Name
        End If
    
    Next
    
    Set vbc = Nothing
    
End Sub
