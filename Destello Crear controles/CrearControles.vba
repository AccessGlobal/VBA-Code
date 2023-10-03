'Módulo estándar de un formulario
Option Compare Database
Option Explicit

Private Sub btnAdd_Click()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-evolucion-de-carga-artesano/
'                      Destello formativo 356
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : btnAdd
' Autor original    : Microsoft | https://learn.microsoft.com/en-us/office/vba/api/access.application.createcontrol
' Creado            : 22/01/2022
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Propósito         : crear un formulario y sus controles en tiempo de ejecución
'---------------------------------------------------------------------------------------------------------------------------------------------
' Objetos           : se incluye toda la acción en el evento "Al hacer clic" de un botón de comando
'---------------------------------------------------------------------------------------------------------------------------------------------
Dim frm As Form
Dim ctrlEtiqueta As Control, ctlTexto As Control, ctrlCombo As Control
Dim intX As Integer
Dim intTextoY As Integer, intComboy As Integer
Dim intEtiquetaX As Integer, intEtiquetaY As Integer
Dim anchoEtiqueta As Long

'Creamos un nuevo formulario
    Set frm = CreateForm
'Podemos aplicar el origen del formulario de esta forma
'        frm.RecordSource = "MiOrigen"
'Podemos acceder a las propiedades del nuevo formulario
'    With frm
'        .Visible = False
'    End With
    
'Posición de los controles
    intX = 200
    intEtiquetaY = 200
    intTextoY = 200
    intComboy = 600
    
'Creamos una etiqueta
    Set ctrlEtiqueta = CreateControl(frm.Name, acLabel, , , "Nueva Etiqueta", intX, intEtiquetaY)

'Según el ancho de etiqueta, colocamos el cuadro de texto
        With ctrlEtiqueta
            anchoEtiqueta = .Width
        End With
        
'Creamos cuadro de texto
    Set ctlTexto = CreateControl(frm.Name, acTextBox, , "", "Texto1", intX + anchoEtiqueta, intTextoY)

'Creamos un combobox
    Set ctrlCombo = CreateControl(frm.Name, acComboBox, , "", "Combo1", intX, intComboy)
    
'Manejamos sus propiedades
        With ctrlCombo
            .RowSource = "SELECT * FROM MiTabla ORDER BY MiOrden"
            .ColumnCount = 2
            .BoundColumn = 1
            .ColumnWidths = "0;6"
        End With
        
 DoCmd.Restore
 
 End Sub