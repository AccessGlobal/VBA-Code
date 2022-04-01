'---------
'MÉTODO 1
'---------
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/drag-drop-en-access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Drag&Drop
' Autor original    : Doug Steele, MVP  AccessHelp@rogers.com
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Actualizado       : mayo 2020
' Propósito         : crear un procedimiento DragDrop que se ejecutará en respuesta a un control que se arrastra a otro control
' ¿Cómo funciona?   : hay dos funciones
'                     DragStart inicializa la funcionalidad "Drag"
'                     DropDetect captura los controles involucrados y sus posiciones
'                     DragStop finaliza la funcionalidad "Drag"
'                     Dispone de tres funciones de apoyo:
'                     - ListBoxExample para realizar la operación entre dos ListBox
'                     - ProcessDrop que permite discernir entre los distintos controles para adaptar el gdato que se coloca (Drop).Esta rutina
'                       debe llamarse desde el evento DetectDrop de cualquier control que desee que pueda ser un destino de un control arrastrado.
' Inputs:               DragForm             Formulario que contiene el control que está siendo arrastrado
'                       DragCtrl             Control, del formulario DragForm, que está siendo arrastrado
'                       DropForm             Formulario que contiene el control donde va a ser colocado el dato
'                       DropCtrl             Control, del formulario DropForm, que va a ser colocado
'                       Button, Shift, X, Y  Parámetros asociados con los eventos del ratón
'                     - ReturnSelectedOption es una función para los OptionGroup
' Más información   : Microsoft Knowledge Base Article 137650
'                     https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/137650
'-----------------------------------------------------------------------------------------------------------------------------------------------

Sub DragStart(SourceFrm As Form)
Dim strIconPath As String

Set DragFrm = SourceFrm
Set DragCtrl = Screen.ActiveControl

CurrentMode = DRAG_MODE
   
'Incorpora el icono de nuestra elección
strIconPath = Application.CurrentProject.Path & "\Recursos\DragDrop.ico"

If Len(Dir$(strIconPath)) > 0 Then
    SetMouseCursorFromFile strIconPath
Else
    SetMouseCursor IDC_IBEAM
End If

End Sub

Sub DragStop()

CurrentMode = DROP_MODE
DropTime = Timer

End Sub

Sub DropDetect(DropFrm As Form, DropCtrl As control, Button As Integer, Shift As Integer, x As Single, y As Single)
    
If CurrentMode <> DROP_MODE Then Exit Sub
    
CurrentMode = NO_MODE

' Se permite el intervalo de temporizador entre el evento MouseUp y el evento
' MouseMove. Esto garantiza que el evento MouseMove no
' invoca el procedimiento de colocación a menos que sea el evento MouseMove
' que Microsoft Access desencadena automáticamente para el control de colocación
' que sigue al evento MouseUp de un control de arrastre. Los eventos
' MouseMove posteriores no superarán la prueba de temporizador y se pasarán por alto.

If Timer - DropTime > MAX_DROP_TIME Then Exit Sub

' ¿Arrastramos o colocamos en nosotros mismos?
If (DragCtrl.Name <> DropCtrl.Name) Or (DragFrm.hwnd <> DropFrm.hwnd) Then
' En caso negativo, se arrastró o colocó correctamente.
    DragDrop DragFrm, DragCtrl, DropFrm, DropCtrl, Button, Shift, x, y
End If

End Sub

Sub DragDrop(DragFrm As Form, DragCtrl As control, _
   DropFrm As Form, DropCtrl As control, _
   Button As Integer, Shift As Integer, _
   x As Single, y As Single)

   ' ¿En qué formulario se colocó?
   ' Es conveniente utilizar el procedimiento DragDrop para
   ' determinar qué operación de arrastrar y colocar se realizó; a continuación, invocar
   ' el código apropiado para tratar los casos especiales.
   Select Case DropFrm.Name
      
      Case "02_DragDropListViews"
         ListBoxExample DragFrm, DragCtrl, DropFrm, DropCtrl, Button, Shift, x, y
      Case Else
         ProcessDrop DragFrm, DragCtrl, DropFrm, DropCtrl, Button, Shift, x, y
   End Select
End Sub

Sub ListBoxExample(DragFrm As Form, DragCtrl As control, DropFrm As Form, DropCtrl As control, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim db As DAO.Database
Dim SQL As String

Set db = CurrentDb()

   ' Crear una instrucción SQL para actualizar el campo Seleccionado del
   ' .. elemento del cuadro de lista de arrastrar/colocado.
SQL = "UPDATE Clientes SET cueact="

   ' Arrastrar desde Lista1 alternar Seleccionado=Verdadero, Lista2 cambia a Falso.
SQL = IIf(DragCtrl.Name = "Lista1", SQL & "False", SQL & "True")
   ' Si no se utiliza la tecla CTRL, modificar únicamente el valor arrastrado.
If (Shift And CTRL_MASK) = 0 Then
    SQL = SQL & " WHERE [Cliente]='" & DragCtrl & "'"
End If

   ' Ejecutar la consulta de actualización para alternar
   ' el campo Seleccionado del registro o los registros de Cliente.
db.Execute SQL

   ' Volver a consultar los controles del cuadro de lista para mostrar
   ' las listas de actualización.
   DragCtrl.Requery
   DropCtrl.Requery

End Sub

Sub ProcessDrop(DragForm As Form, DragCtrl As control, DropForm As Form, DropCtrl As control, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ctlCurr As control
Dim strSelectedItems As String
Dim varCurrItem As Variant

On Error GoTo linErr

If TypeOf DragCtrl Is CheckBox Then
    DropCtrl = IIf(DragCtrl, "True", "False")
ElseIf TypeOf DragCtrl Is OptionGroup Then
'Adaptamos los datos a los controles de origen y de destino
    If TypeOf DropCtrl Is TextBox Then
        DropCtrl = DragCtrl & ": " & ReturnSelectedOption(DragCtrl)
    End If
Else
       DropCtrl = DragCtrl
End If

Exit Sub

linErr:
'Desarrolla el control de errores que más te guste
End Sub

Function ReturnSelectedOption(OptionGroup As OptionGroup) As String
Dim ctlCurr As control
Dim booGetText As Boolean
Dim strSelected As String

On Error GoTo linErr

For Each ctlCurr In OptionGroup.Controls
    If TypeOf ctlCurr Is OptionButton Or TypeOf ctlCurr Is CheckBox Then
        If ctlCurr.OptionValue = OptionGroup.Value Then
            strSelected = ctlCurr.Name
            booGetText = True
            Exit For
        End If
    ElseIf TypeOf ctlCurr Is ToggleButton Then
        If ctlCurr.OptionValue = OptionGroup.Value Then
            ReturnSelectedOption = ctlCurr.Caption
            booGetText = False
            Exit For
        End If
    End If
Next ctlCurr

If booGetText Then
    For Each ctlCurr In OptionGroup.Controls
        If TypeOf ctlCurr Is Label Then
            If ctlCurr.Parent.ControlName = strSelected Then
                ReturnSelectedOption = ctlCurr.Caption
                Exit For
            End If
        End If
    Next ctlCurr
End If

Exit Function

linErr:
'Desarrolla el control de errores que más te guste
End Function



