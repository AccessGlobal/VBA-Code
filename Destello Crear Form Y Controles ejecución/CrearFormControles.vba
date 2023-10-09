'Módulo estándar 
Option Compare Database
Option Explicit

Dim strSQL As String

Function creaCuadrante(IdProyecto As Long, fIni As Date, fFin As Date) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crea-formulario-en-tiempo-de-ejecucion
'                     Destello formativo 359
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : creaCuadrante
' Autor             : Alba Salvá | albasalva@access-global.net
' Fecha             : marzo 2013
' Propósito         : crear un formulario (cuadrante) y sus controles en tiempo de ejecución, desde una consulta SQL
' Retorno           : 1/0 según haya tenido éxito o no la creación del formulario
' Argumento/s       : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     idproyecto      Obligatorio      id del proyecto que se está analizando
'                     fIni            Obligatorio      Fecha de inicio
'                     fFin            Obligatorio      Fecha de finalización
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : al tratarse de código extraído de una aplicación que actualmente se encuentra en funcionamiento, no se puede recrear
'                     la consulta a menos que se creen las tablas y campos necesarios para ello.
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim dtTemp
Dim ctlLeft       As Long
Dim ctlWidth      As Long
Dim ctlSep        As Long
Dim lblTop        As Long
Dim lblHeight     As Long
Dim lblWidth      As Long
Dim txtTop        As Long
Dim txtHeight     As Long
Dim tbxWidth      As Long
Dim ctlDest       As Control
Dim intCont       As Integer
Dim strFORM       As String
Dim objFormatCond As FormatCondition
Dim frm           As Form
    
    DoEvents
   
    On Error Resume Next
'Comprueba si el formulario está cargado en este momento y si lo está, lo cierra
    If CurrentProject.AllForms("tmpCuadrante").IsLoaded Then
        DoCmd.Close acForm, "tmpCuadrante", acSaveNo
    End If
'Elimina el formulario para crearlo de nuevo
    DoCmd.DeleteObject acForm, "frmCuadrante"
    
    On Error GoTo 0
'Carga los datos de la consulta en una tabla temporalq ue borramos inicialmente
    CurrentDb.Execute "DELETE FROM tbtRegTiempo"
    
    strSQL = "INSERT INTO tbtRegTiempo (IdProyecto, IdEmp, dtFecha, sngHorasT, sngHorasF, strNotas) " & vbCrLf & _
        "SELECT IdProyecto, IdEmp, dtFecha, sngHorasT, sngHorasF, strNotas " & vbCrLf & _
        "FROM tblRegTiempo " & vbCrLf & _
        "WHERE IdProyecto=" & IdProyecto
             
    CurrentDb.Execute strSQL
'Vuelve a cargar los datos que se precisan
    DoCmd.OpenQuery "qTrcRegTiempoF"
    DoCmd.Close acQuery, "qTrcRegTiempoF", acSaveYes
    
'Establece posición y tamaño de los controles que previamente habían sido grabados en el registro (destello 357)
    lblTop = GetSetting("ViTools", "Generator", "lblTop", -60)
    lblHeight = GetSetting("ViTools", "Generator", "lblHeight", -655)
    lblWidth = GetSetting("ViTools", "Generator", "lblWidth", lblWidth)
    
    txtTop = GetSetting("ViTools", "Generator", "tbxTop", -57)
    txtHeight = GetSetting("ViTools", "Generator", "tbxHeight", -255)
    tbxWidth = GetSetting("ViTools", "Generator", "tbxWidth", tbxWidth)
    
    ctlSep = GetSetting("ViTools", "Generator", "ctlSep", ctlSep)
    ctlWidth = GetSetting("ViTools", "Generator", "ctlWidth", ctlWidth)
    ctlLeft = GetSetting("ViTools", "Generator", "ctlLeft", ctlLeft)
   
    On Error GoTo 0
    
'Crea el una copia del formulario que se utiliza como plantilla
    DoCmd.CopyObject , "frmCuadrante", acForm, "tmpCuadrante"
'Muestra en oculto el formulario para incorporarle los nuevos controles
    DoCmd.OpenForm ("frmCuadrante"), View:=acDesign, WindowMode:=acHidden

    intCont = 1
    strFORM = "frmCuadrante"
    
    On Error GoTo lbError
'Crea los controles y establece el formato condicional de alguno de ellos (destello 358)
    For dtTemp = fIni To fFin
        DoEvents
        
        Set ctlDest = CreateControl(strFORM, acLabel, acHeader, , , ctlLeft, lblTop, ctlWidth, lblHeight)
        ctlDest.name = "lblFecha" & intCont
        ctlDest.FontSize = 9
        
        ctlDest.BackStyle = 1
        ctlDest.ForeColor = vbBlack
        
        Select Case Weekday(dtTemp)
            Case vbMonday To vbFriday
                ctlDest.BackStyle = 0
            Case vbSaturday
                ctlDest.BackColor = clrYellow
            Case vbSunday
                ctlDest.ForeColor = clrWhite
                ctlDest.BackColor = clrRed
        End Select
        ctlDest.Caption = Replace(dtTemp, "/", vbCrLf)
        ctlDest.TextAlign = 2 'Center
        
        Set ctlDest = CreateControl(strFORM, acTextBox, acDetail, , , ctlLeft, txtTop, ctlWidth, txtHeight)
        ctlDest.name = "txtFecha" & intCont
        ctlDest.ControlSource = CStr(dtTemp)
        ctlDest.TextAlign = 2
        ctlDest.FontSize = 9
        
        Set objFormatCond = ctlDest.FormatConditions.Add(acExpression, , "[Tipo] ='F'")
        With objFormatCond
            .FontBold = True
            .BackColor = vbBlue
            .ForeColor = vbWhite
        End With
        ctlLeft = ctlLeft + ctlSep
        
        intCont = intCont + 1
    Next
    
'Adapta el tamaño del formulario
    Forms(strFORM).InsideWidth = Forms(strFORM).InsideWidth + (ctlSep - ctlWidth)
    
    DoCmd.Close acForm, "frmCuadrante", acSaveYes

    DoCmd.Restore
    
    Application.SetHiddenAttribute acForm, "frmCuadrante", True
    
    creaCuadrante = 1
    
    GoTo lbFinally
    
lbError:
    If Err = 2100 Then
        Select Case MsgBox("No se puede mostrar el resultado dentro de Access," & vbCrLf & _
                "¿Desea crearlo en Excel?", vbQuestion Or vbYesNo)
            Case vbYes
                DoCmd.Close acForm, "frmCuadrante", acSaveNo
                DoCmd.DeleteObject acForm, "frmCuadrante"
                Call CreaExcel ' (Función externa)
                DoCmd.Close acForm, "frmCuadrante", acSaveYes
                creaCuadrante = 2
            Case Else
                creaCuadrante = 0
        End Select
    Else
        MsgBox "Error: " & Err & vbCrLf & Err.Description
        creaCuadrante = 0
    End If
    
lbFinally:

End Function