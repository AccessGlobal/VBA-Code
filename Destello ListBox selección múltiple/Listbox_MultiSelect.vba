'Módulo de un formulario
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-cuadro-de-lista-con-seleccion-multiple/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Guardar_Click
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : junio 2010
' Propósito         : carga del file picker a la apertura del formulario
' Argumento/s       : No se dispone de argumentos. Se ejecuta el la acción en la apertura de un formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fdialog As Office.FileDialog
'Dim o_fso As New FileSystemObject
Dim varfile As Variant
Dim ruta As String

Me.ListaDocs.RowSource = ""
    
Set fdialog = Application.FileDialog(msoFileDialogFilePicker)
    With fdialog
      .AllowMultiSelect = True
      .Title = "Selecciona los archivos que desees"
      .filters.Clear
      .filters.Add "All Files", "*.*"

        If .Show = True Then
            For Each varfile In .SelectedItems
                ruta = CStr(varfile)
'                Set Archivo = o_fso.GetFile(varfile)
                    Me.ListaDocs.AddItem ruta
'                Set Archivo = Nothing
            Next
        End If
    End With
Set fdialog = Nothing

End Sub

Private Sub Guardar_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-cuadro-de-lista-con-seleccion-multiple/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Guardar_Click
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : junio 2010
' Propósito         : elementos necesarios para manejar un ListBox de selección múltiple
' Argumento/s       : No se dispone de argumentos. Se ejecuta el la acción click de un botón de comando
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim ArrayDir As Variant
Dim J As Long, sizefil As Long
Dim intWhere As Integer, contador As Integer
Dim subdir As String
Dim fso, objfile
Dim oItem As Variant
Dim stemp As String

If Me.ListaDocs.ListIndex = -1 Then Exit Sub
   
For Each oItem In Me.ListaDocs.ItemsSelected
    contador = contador + 1
    stemp = Me.ListaDocs.ItemData(oItem)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        Set objfile = fso.GetFile(stemp)
            sizefil = objfile.Size
        Set objfile = Nothing
    Set fso = Nothing
       
    ArrayDir = Split(stemp, "\")
    J = UBound(ArrayDir) - LBound(ArrayDir) + 1
    
    subdir = ArrayDir(J - 1) 'Nombre del fichero
    
    intWhere = 0
    intWhere = intWhere + InStr(1, subdir, "'")
    intWhere = intWhere + InStr(1, subdir, """")
    intWhere = intWhere + InStr(1, subdir, "/")
    intWhere = intWhere + InStr(1, subdir, "\")
    
    If intWhere <> 0 Then
        MsgBox "Se han utilizado caracteres no permitidos"
    End If
    
    Debug.Print subdir
Next

End Sub

Private Sub Salir_Click()

DoCmd.Close acForm, "FormLista"

End Sub
