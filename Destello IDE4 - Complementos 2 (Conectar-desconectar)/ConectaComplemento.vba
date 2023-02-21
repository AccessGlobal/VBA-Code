'Colocar el siguiente código en los eventos "Al hacer click" de dos botones
Private Sub btnConectar_Click()
Dim strAddin As String
Dim intStr As Integer
Dim intI As Integer

    If IsNull(Me.lstAddIns.Value) Then
        MsgBox "Debes seleccionar un complemento de la lista. Selecciónalo y luego podrás conectarlo"
        Exit Sub
    End If

    strAddin = Me.lstAddIns.Value
    
    intStr = InStr(1, strAddin, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strAddin = strAddin
    Else
        strAddin = Left(strAddin, intStr - 1)
    End If
    
    If strAddin = "" Then
        Exit Sub
    End If
    
    If Application.VBE.AddIns(strAddin).Connect = True Then
        MsgBox "El complemento ya está conectado"
    Else
        Conectar (strAddin)
        ListaComplementos Me
    End If
    
End Sub

Private Sub btnDesconectar_Click()
Dim strAddin As String
Dim intStr As Integer
Dim intI As Integer

    If IsNull(Me.lstAddIns.Value) Then
        MsgBox "Debes seleccionar un complemento de la lista. Selecciónalo y luego podrás desconectarlo"
        Exit Sub
    End If
    
    strAddin = Me.lstAddIns.Value
    
    intStr = InStr(1, strAddin, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strAddin = strAddin
    Else
        strAddin = Left(strAddin, intStr - 1)
    End If
    
    If strAddin = "" Then
        Exit Sub
    End If
    
    If Application.VBE.AddIns(strAddin).Connect = False Then
        MsgBox "El complemento no está conectado"
    Else
        Desconectar (strAddin)
        ListaComplementos Me

    End If
    
End Sub

'Colocar código en módulo estándar
Option Compare Database
Option Explicit

Public Sub ListaComplementos(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net\vbide-series-complementos-conectar-y-desconectar-complementos
'                     Destello formativo 270
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConectaComplementos
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : mostrar un listado con todos los complementos del VBIDE de nuestro programa para conectarlos y desconectarlos
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
            
            If AddIns(intI).Connect = True Then
                frm.lstAddIns.AddItem AddinName & " | Conectado"
            Else
                frm.lstAddIns.AddItem AddinName & " | Desconectado"
            End If
            
        Next
  
    Set AddIns = Nothing

End Sub

Public Sub Conectar(strAddin As String)

    Application.VBE.AddIns(strAddin).Connect = True

End Sub

Public Sub Desconectar(strAddin As String)

    Application.VBE.AddIns(strAddin).Connect = False

End Sub
