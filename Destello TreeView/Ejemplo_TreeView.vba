(Formulario Drag&Drop)
Option Compare Database
Option Explicit

Private Sub Cerrar_Click()

    DoCmd.Close acForm, "TreeViewDragDrop"
    
    Set tv = Nothing
    
End Sub

Private Sub CollapseAll_Click()
Dim I As Integer
Dim objnode

    On Error Resume Next
    
    TV1.SetFocus
    
    For I = 1 To TV1.Nodes.Count - 1
        TV1.Nodes(I).Expanded = False
        Err.Clear
    Next I
    
End Sub

Private Sub ExpandAll_Click()
Dim I As Integer
Dim objnode

    On Error Resume Next
    
    TV1.SetFocus
    
'Recorre los nodos y los expande mediante la porpiedad "Expanded"
    For I = 1 To TV1.Nodes.Count - 1
        TV1.Nodes(I).Expanded = True
        Err.Clear
    Next I
    
End Sub
Private Sub Form_Open(Cancel As Integer)

Call CreaTreeViewProductos

On Error Resume Next

For I = 1 To TV1.Nodes.Count - 1
    TV1.Nodes(I).Expanded = True
    Err.Clear
Next I

Set objnode = TV1.Nodes(1)
    With objnode
        .Selected = True
        Err.Clear
    End With
Set objnode = Nothing

Set tv = Me.TV1.Object

Set imgListObj = Me.TV1ImageList.Object
tv.ImageList = imgListObj

End Sub

Private Function CreaTreeViewProductos()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/drag-drop-en-access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CreaTreeViewProductos
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : febrero 2018
' Propósito         : crear un treeview con los tipos (elementos de la tabla "tipo")
' Retorno           : sin retorno
' Argumento/s       : no precisa ningún argumento
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en el evento de carga de cualquier formulario que desees.
'
'                     Private Sub Form_load()
'
'                         Call CreaTreeViewProductos
'
'                      End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim tipo2 As String

tipo2 = "Productos"
lngGrey = rgb(150, 150, 150)

With Me.TV1ImageList
    With .ListImages
        .Clear
        .Add Key:="ImgProductos", Picture:=LoadPicture(CurrentProject.Path & "\Galería\Productos.bmp")
    End With
End With

With Me.TV1
    .Nodes.Clear
    .style = tvwTreelinesPlusMinusPictureText
    .LineStyle = tvwRootLines
    .Indentation = 240
    .Appearance = ccFlat
    .HideSelection = False
    .BorderStyle = ccFixedSingle
    .HotTracking = True
    .FullRowSelect = False
    .CheckBoxes = False
    .SingleSel = False
    .Sorted = False
    .Scroll = True
    .LabelEdit = tvwManual
    .Font.Name = "Century Gothic"
    .Font.Size = 10
    .ImageList = Me.TV1ImageList.Object
End With

'Añadimos el root productos
TV1.Nodes.Clear
nodeKey = "n0"
Set objnode = TV1.Nodes.Add(, , nodeKey, tipo2)
    objnode.ForeColor = lngGrey
    objnode.Image = "ImgProductos"
Set objnode = Nothing

'Añadimos todos los nodos
Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM tipo WHERE tipocod=2 ORDER BY idtipo ASC")
    Do Until rstTable.EOF
        nodeKey = "n" & rstTable!idtipo
        If Not IsNull(rstTable!tipopadre) Then
            parentKey = "n" & rstTable!tipopadre
        Else
            parentKey = "n0"
        End If
        Set objnode = TV1.Nodes.Add(parentKey, tvwChild, nodeKey, rstTable!tiponom)
            objnode.ForeColor = lngGrey
            objnode.Image = "ImgProductos"
        Set objnode = Nothing
    rstTable.MoveNext
    Loop
Set rstTable = Nothing

End Function

Private Sub TV1_NodeClick(ByVal Node As Object)
Dim SelectionNode As MSComctlLib.Node
    
'Ensure that the clicked node equals the selected node in the tree
If Not Node Is Nothing Then
    Set SelectionNode = Node
       If SelectionNode.Expanded = True Then
            SelectionNode.Expanded = False
        Else
            SelectionNode.Expanded = True
        End If
End If

End Sub

Private Sub TV1_OLEStartDrag(Data As Object, AllowedEffects As Long)
    
    Set Me.TV1.SelectedItem = Nothing

End Sub

Private Sub TV1_OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim SelectedNode As MSComctlLib.Node
Dim nodOver As MSComctlLib.Node
    
If tv.SelectedItem Is Nothing Then
'Selecciona un nodo si no hay uno seleccionado
    Set SelectedNode = tv.HitTest(x, y)
    If Not SelectedNode Is Nothing Then
        SelectedNode.Selected = True
    End If
Else
    If tv.HitTest(x, y) Is Nothing Then
'En este sitio puedes poner la función que quieras
    Else
'Marca el nodo sobre el que se posiciona el ratón
        Set nodOver = tv.HitTest(x, y)
        Set tv.DropHighlight = nodOver
    End If
End If

End Sub

Private Sub TV1_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sourceNode As MSComctlLib.Node
Dim SourceParentNode As MSComctlLib.Node
Dim targetNode As MSComctlLib.Node
    
Dim tmpRootNode As MSComctlLib.Node
Dim strtmpNodKey As String
Dim ChildNode As MSComctlLib.Node
    
Dim strSPKey As String
Dim strTargetKey As String
    
Dim strsQL As String
Dim intKey As Integer
Dim intPKey As Integer
    
    
    Set sourceNode = tv.SelectedItem
        
    Set SourceParentNode = sourceNode.Parent
    Set targetNode = tv.HitTest(x, y)
                
    On Error GoTo LinError
        
    If SourceParentNode Is Nothing Then
        strSPKey = "Empty"
    Else
        strSPKey = SourceParentNode.Key
    End If
        
    Select Case True
        Case targetNode Is Nothing
            strTargetKey = "Empty"
        Case targetNode.Key = ""
            strTargetKey = "Empty"
            Set targetNode = Nothing
        Case Else
            strTargetKey = targetNode.Key
    End Select
        
    If strTargetKey = strSPKey Then Exit Sub
           
    Set sourceNode.Parent = targetNode
               
    If targetNode Is Nothing Then
        intKey = Val(Mid(sourceNode.Key, 2))
        strsQL = "UPDATE tipo SET tipopadre = Null WHERE idtipo = " & intKey
    Else
        intKey = Val(Mid(sourceNode.Key, 2))
        intPKey = Val(Mid(targetNode.Key, 2))
                
        strsQL = "UPDATE tipo SET tipopadre = " & intPKey & " WHERE idtipo = " & intKey
    End If
            
    'Modifica la tabla  con el nuevo cambio de arrastrar
    CurrentDb.Execute strsQL, dbFailOnError
            
    If sourceNode.Parent Is Nothing Then
        sourceNode.Root.Sorted = True
    Else
        sourceNode.Parent.Sorted = True
    End If
                
    tv.Nodes(sourceNode.Key).Selected = True
    
    Exit Sub
    
LinError:
    'Crea el control de errores que más te guste
    CreaTreeViewProductos 'Refresca el TreeView con los datos iniciales
    
End Sub

Private Sub TV1_OLECompleteDrag(Effect As Long)

    Set tv.DropHighlight = Nothing

End Sub


(Formulario TreeView ejemplo)
Option Compare Database
Option Explicit

Private Sub Cerrar_Click()

    DoCmd.Close acForm, "TreeViewEjemplo"

End Sub

Private Sub CollapseAll_Click()
Dim I As Integer
Dim objnode

    On Error Resume Next
    
    TV1.SetFocus
    
    For I = 1 To TV1.Nodes.Count - 1
        TV1.Nodes(I).Expanded = False
        Err.Clear
    Next I
    
End Sub

Private Sub ExpandAll_Click()
Dim I As Integer
Dim objnode

    On Error Resume Next
    
    TV1.SetFocus
    
'Recorre los nodos y los expande mediante la porpiedad "Expanded"
    For I = 1 To TV1.Nodes.Count - 1
        TV1.Nodes(I).Expanded = True
        Err.Clear
    Next I
    
End Sub

Private Sub TV1_NodeClick(ByVal Node As Object)
Dim objnode As Node
Dim Nom As String

'Capturamos el nombre del nodo
    Nom = TV1.SelectedItem

    Debug.Print Nom
'Cambiamos la imagen del nodo
    Set objnode = TV1.SelectedItem
        objnode.Image = "OpenFolder"
        Debug.Print objnode.FirstSibling
        Debug.Print objnode.Root
        objnode.Bold = True
        objnode.BackColor = vbGreen
    Set objnode = Nothing

'Podemos incluir cualquier función que queramos
    
       
End Sub

Private Sub Form_Load()

    Me.Lite1 = "TreeView series: Nodos"

    With Me.TVImageList
        With .ListImages
            .Clear
            .Add Key:="OpenFolder", Picture:=LoadPicture(CurrentProject.Path & "\Galería\OpenFolder.bmp")
            .Add Key:="ClosedFolder", Picture:=LoadPicture(CurrentProject.Path & "\Galería\ClosedFolder.bmp")
            .Add Key:="File", Picture:=LoadPicture(CurrentProject.Path & "\Galería\File.bmp")
        End With
    End With

    With Me.TV1
'Limpia los nodos
        .Nodes.Clear
'Apariencia: ccFlat | cc3D
        .Appearance = ccFlat
'Estilo del borde: ccNone | ccFixedSingle
        .BorderStyle = ccFixedSingle
'Incluye o no objetos checkbox
        .CheckBoxes = False
'Activado o desactivado
        .Enabled = True
'Tipo de letra
        .Font.Name = "Century Gothic"
'Tamaño de letra
        .Font.Size = 9
'Selección de fila completa: indica si el resalte abarca al ancho de TreeView
        .FullRowSelect = False
'Altura
'    .Height
'Selección oculta: Obtiene o establece un valor que indica si el nodo seleccionado permanece resaltado incluso cuando el objeto ha perdido el foco.
        .HideSelection = False
'Indica si los nodos proporcionan comentarios cuando el mouse se mueve sobre ellos
        .HotTracking = True
'Sangía: ancho de sangría de los nodos, en píxeles
        .Indentation = 570
'Edición de etiquetas: dos opciones tvwAutomatic | tvwManual
        .LabelEdit = tvwManual
'Estilo de líneas: tvwTreeLines | tvwRootLines
        .LineStyle = tvwRootLines
'Icono del ratón: ccDefault | ccArrow | ccCross | ccIBeam | ccIcon | ccSize | ccSizeNESW | ccSizeNS | ccSizeNWSE | ccSizeEW | ccUpArrow | ccHourglass | ccNoDrop | ccArrowHourglass | ccArrowQuestion | ccSizeAll | ccCustom
'    .MousePointer
'    .PathSeparator
'Desplazamiento: sí o no
        .Scroll = True
'Selección única
        .SingleSel = False
'Ordenación:Cuando se establece en falso (predeterminado), los nodos se mostrarán en el orden en que se agregaron a la matriz .Nodes. Cuando se establece en verdadero, los nodos se ordenarán alfabéticamente.
        .Sorted = True
'Estilo: tvwTextOnly | tvwPictureText | tvwPlusMinusText | tvwPlusPictureText | tvwTreelinesText | tvwTreelinesPictureText | tvwTreelinesPlusMinusText | tvwTreelinesPlusMinusPictureText )
        .style = tvwTreelinesPlusMinusPictureText
'    .Width 'Ancho
        .ImageList = Me.TVImageList.Object
    End With

    Call createTree

End Sub

Private Sub createTree()
Dim nomold As String

    On Error Resume Next
    
    lngGrey = rgb(150, 150, 150)
    
    Keynod = "1"
    Keynod1 = "1"
    Keynod2 = "1"
    
    Set rstTable = CurrentDb.OpenRecordset("ConsultaEjemploTreeView")
        Do Until rstTable.EOF
            If nomold = rstTable!doctree1nom Then GoTo LinNext
            Set objnode = TV1.Nodes.Add(, tvwChild, "A" + Keynod, rstTable!doctree1nom)
                objnode.Selected = False
                objnode.ForeColor = lngGrey
                objnode.Image = "ClosedFolder"
                Call SubNodes1(rstTable!doctree1nom, Keynod)
                Keynod = str(CInt(Keynod) + 1)
                nomold = rstTable!doctree1nom
                Err.Clear
LinNext:
        rstTable.MoveNext
        Loop
        Set objnode = Nothing
    Set rstTable = Nothing
    
    Set objnode = TV1.Nodes(1)
        With objnode
            .Selected = True
            Err.Clear
        End With
    Set objnode = Nothing

End Sub

Private Sub SubNodes1(ByVal F1 As String, ByVal Keynod As String)
Dim objnode As Node

    Set rstTable1 = CurrentDb.OpenRecordset("SELECT * FROM ConsultaEjemploTreeView WHERE doctree1nom like '" & F1 & "'")
        Do Until rstTable1.EOF
            If Not IsNull(rstTable1!iddoctree2) Then
                Set objnode = TV1.Nodes.Add("A" + Keynod, tvwChild, "B" + Keynod1, rstTable1!doctree2nom)
                objnode.Selected = False
                objnode.ForeColor = lngGrey
                objnode.Image = "ClosedFolder"
                Keynod1 = str(CInt(Keynod1) + 1)
            End If
        rstTable1.MoveNext
        Loop
    Set rstTable1 = Nothing

End Sub

Private Sub TV1_OLEStartDrag(Data As Object, AllowedEffects As Long)
    
    Set Me.TV1.SelectedItem = Nothing

End Sub

Private Sub TV1_OLEDragOver(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim SelectedNode As MSComctlLib.Node
Dim nodOver As MSComctlLib.Node
    
If tv.SelectedItem Is Nothing Then
        'Select a node if one is not selected
    Set SelectedNode = tv.HitTest(x, y)
    If Not SelectedNode Is Nothing Then
        SelectedNode.Selected = True
    End If
Else
    If tv.HitTest(x, y) Is Nothing Then
        'do nothing
    Else
            'Highlight the node the mouse is over
        Set nodOver = tv.HitTest(x, y)
        Set tv.DropHighlight = nodOver
    End If
End If

End Sub

Private Sub TV1_OLEDragDrop(Data As Object, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sourceNode As MSComctlLib.Node
Dim SourceParentNode As MSComctlLib.Node
Dim targetNode As MSComctlLib.Node
    
Dim tmpRootNode As MSComctlLib.Node
Dim strtmpNodKey As String
Dim ChildNode As MSComctlLib.Node
    
Dim strSPKey As String
Dim strTargetKey As String
    
Dim strsQL As String
Dim intKey As Integer
Dim intPKey As Integer
    
    
Set sourceNode = tv.SelectedItem
    
Set SourceParentNode = sourceNode.Parent
Set targetNode = tv.HitTest(x, y)
            
On Error GoTo LinError
    
If SourceParentNode Is Nothing Then
    strSPKey = "Empty"
Else
    strSPKey = SourceParentNode.Key
End If
    
Select Case True
    Case targetNode Is Nothing
        strTargetKey = "Empty"
    Case targetNode.Key = ""
        strTargetKey = "Empty"
        Set targetNode = Nothing
    Case Else
        strTargetKey = targetNode.Key
End Select
    
If strTargetKey = strSPKey Then Exit Sub
       
Set sourceNode.Parent = targetNode
           
If targetNode Is Nothing Then
    intKey = Val(Mid(sourceNode.Key, 2))
    strsQL = "UPDATE tipo SET tipopadre = Null WHERE idtipo = " & intKey
Else
    intKey = Val(Mid(sourceNode.Key, 2))
    intPKey = Val(Mid(targetNode.Key, 2))
            
    strsQL = "UPDATE tipo SET tipopadre = " & intPKey & " WHERE idtipo = " & intKey
End If
        
'Modifica la tabla  con el nuevo cambio de arrastrar
CurrentDb.Execute strsQL, dbFailOnError
        
If sourceNode.Parent Is Nothing Then
    sourceNode.Root.Sorted = True
Else
    sourceNode.Parent.Sorted = True
End If
            
tv.Nodes(sourceNode.Key).Selected = True

Exit Sub

LinError:
'Crea el control de errores que más te guste

End Sub

Private Sub TV1_OLECompleteDrag(Effect As Long)

    Set tv.DropHighlight = Nothing

End Sub




