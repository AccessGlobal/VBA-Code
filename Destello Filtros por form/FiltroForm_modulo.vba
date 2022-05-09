'Este cloque de código es a nivel de módulo
Option Compare Database
Option Explicit

Sub CrearFiltros()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-gestionamos-el-filtro-por-formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CrearFiltros
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 22
' Propósito         : crear un menú contextual
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/office/vba/api/access.application.commandbars
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
'
'        Call CrearFiltros
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim cmdMenu As Office.CommandBar
Dim cmdSubmenu As CommandBarControl
Dim NewControl As CommandBarButton

On Error Resume Next

Set cmdMenu = CommandBars.Add("Filtros de formulario", msoBarPopup, False, True)

    Set NewControl = cmdMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
        With NewControl
            .Caption = "Filtrar por formulario"
            .OnAction = "=FiltroFormTest()"
            .BeginGroup = True
            .FaceId = 327
            .Style = msoButtonIconAndCaption
        End With
    Set NewControl = Nothing
    
    Set NewControl = cmdMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
        With NewControl
            .Caption = "Quitar filtro"
            .OnAction = "=FiltroFormTest()"
            .BeginGroup = True
            .FaceId = 327
            .Style = msoButtonIconAndCaption
        End With
    Set NewControl = Nothing
    
Set cmdMenu = Nothing
    
End Sub

Function FiltroFormTest()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-gestionamos-el-filtro-por-formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FiltroFormTest
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 22
' Propósito         : selecciona la acción dependiendo del cpation del control que realiza la llamada
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim strTest As String
Dim ctrl As CommandBarControl

Set ctrl = CommandBars("Filtros de formulario").Controls(1)
  strTest = ctrl.Caption
  Select Case strTest
    Case "Filtrar por pormulario"
      DoCmd.RunCommand acCmdFilterByForm
      DoCmd.RunCommand acCmdClearGrid
      ctrl.Caption = "Aplicar filtro"
      CommandBars("Filtros de formulario").Controls(2).Visible = True
    
    Case "Aplicar filtro"
      DoCmd.RunCommand acCmdApplyFilterSort
      ctrl.Caption = "Quitar filtro"
      CommandBars("Filtros de formulario").Controls(2).Visible = False
    
    Case "Quitar filtro"
      DoCmd.RunCommand acCmdRemoveFilterSort
      ctrl.Caption = "Filtrar por pormulario"
  End Select

End Function