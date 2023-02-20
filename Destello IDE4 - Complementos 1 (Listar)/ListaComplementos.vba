'Codigo del formulario
Private Sub Form_Load()
    
    ListaComplementos Me

End Sub

Private Sub lstAddIns_DblClick(Cancel As Integer)
Dim strAddIn As String
Dim intStr As Integer
Dim intI As Integer
Dim AddIns As VBIDE.AddIns
Dim AddinName As String, AddInDescription As String
Dim AddInGuid As String
Dim mensaje As String

    strAddIn = Me.lstAddIns.Value
    
    intStr = InStr(1, strAddIn, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strAddIn = strAddIn
    Else
        strAddIn = Left(strAddIn, intStr - 1)
    End If
    
    If strAddIn = "" Then
        Exit Sub
    End If
       
'Como no es una colección, no podemos recorrerlo con For Each...Next hay qye hacerlo con For
    Set AddIns = Application.VBE.AddIns
    
        For intI = 1 To AddIns.Count
            AddinName = AddIns(intI).ProgId
            If AddinName = strAddIn Then
                AddInDescription = AddIns(intI).Description
                AddInGuid = AddIns(intI).Guid
                mensaje = "AddIn: " & AddinName & _
                        vbNewLine & _
                        "Descripción: " & AddInDescription & _
                        vbNewLine & _
                        "GUID: " & AddInGuid
                MsgBox mensaje, vbInformation + vbOKOnly, "Detalle del complemento"
                Exit Sub
            End If
            
        Next
  
    Set AddIns = Nothing

End Sub


'Código de un módulo estándar
Option Compare Database
Option Explicit

Public Sub ListaComplementos(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-complementos-listar-complementos/
'                     Destello formativo 269
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListaComplementos
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : mostrar un listado con todos los complementos del VBIDE de nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : No se mostrarán los complementos de Access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta las líneas y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ListaComplementos Me
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim intI As Integer
Dim AddIns As VBIDE.AddIns
Dim AddinName As String

'Limpiamos el listbox
    frm.lstAddIns.RowSource = ""

'Como no es una colección, no podemos recorrerlo con For Each...Next hay qye hacerlo con For
    Set AddIns = Application.VBE.AddIns
    
        For intI = 1 To AddIns.Count
        
            AddinName = AddIns(intI).ProgId
            
            frm.lstAddIns.AddItem AddinName
            
        Next
  
    Set AddIns = Nothing

End Sub