Option Compare Database
Option Explicit

Sub ListadoReferencias(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-listar-referencias/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoReferencias
' Autor original    : Alba Salvá
' Creado            : Alba Salvá
' Adaptado por      : Luis Viadel
' Propósito         : mostrar un listado de referencias de nuestro programa con su descripción, tipo, ubicación, su dirección física y el GUID
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/es-es/office/vba/api/access.application.vbe
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    listadoReferencias Me
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim objRef As VBIDE.reference   ' Referencia
Dim objRefs As VBIDE.References ' Colección

    frm.lstRefs.RowSource = ""
    frm.lstRefs.Requery
 
'Listar referencias utilizando VBIDE
    Set objRefs = Application.VBE.ActiveVBProject.References
        For Each objRef In objRefs
            frm.lstRefs.AddItem objRef.name & " | " & objRef.Description & " | " & IIf(objRef.BuiltIn, "Interna", "Externa")
'            Debug.Print objRef.name, objRef.Description
'            Debug.Print Space(14) & objRef.Guid
'            Debug.Print Space(14) & objRef.BuiltIn
'            Debug.Print Space(14) & objRef.IsBroken
'            Debug.Print Space(14) & objRef.Type
'            Debug.Print Space(14) & objRef.Major
'            Debug.Print Space(14) & objRef.Minor
'            Debug.Print vbNewLine
        Next objRef
        Form_Referencias.RefTotal = objRefs.Count
    Set objRefs = Nothing

End Sub

Sub ListadoReferenciasAccess(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-listar-referencias/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListadoReferenciasAccess
' Autor original    : Luis Viadel
' Creado            : desconocido
' Propósito         : mostrar un listado de referencias de nuestro programa con sus propiedades
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/api/access.reference
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ListadoReferenciasAccess Me
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim objRef As reference

    frm.lstRefs.RowSource = ""
    frm.lstRefs.Requery
 
'Listar referencias mediante el objeto "References"
    Set objRef = References!Access
        For Each objRef In References
            frm.lstRefs.AddItem objRef.name
'            Debug.Print objRef.name
'            Debug.Print Space(14) & objRef.FullPath
'            Debug.Print Space(14) & objRef.Guid
'            Debug.Print Space(14) & objRef.BuiltIn
'            Debug.Print Space(14) & objRef.IsBroken
'            Debug.Print Space(14) & objRef.Kind
'            Debug.Print Space(14) & objRef.Major
'            Debug.Print Space(14) & objRef.Minor
'            Debug.Print vbNewLine
        Next objRef

    Set objRef = Nothing

End Sub