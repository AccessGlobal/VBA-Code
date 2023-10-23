'Módulo estándar: "modMacros"
Option Compare Database
Option Explicit

Dim dicMacros As Dictionary

Public Sub ConversorMacros()
'---------------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-convertir-macros-a-vba-sin-asistente/
'                     Destello 366
'---------------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ConversorMacros
' Autor original    : Alba Salvá
' Creado            : octubre 2023
' Propósito         : pasar todas las macros a código VBA
'---------------------------------------------------------------------------------------------------------------------------------------------------------
Dim oFrm    As Object
Dim frm     As Access.Form
Dim oRpt    As Object
Dim rpt     As Access.Report
Dim ctl     As Access.Control
Dim prp     As DAO.Property
Dim oScr    As Object
    
Dim nomFile As String
Dim ruta    As String
Dim vDic    As Variant
    
Dim iNumForms As Integer _
    , iNumReports As Integer _
    , iNumMacros As Integer
Dim sMsg    As String

Dim cForms As Integer, cReports As Integer
    
    On Error GoTo Error_Handler
    
    sMsg = "Antes de ejecutar este código," & vbCrLf & _
        "haga una copia de seguridad de la base de datos" & vbCrLf & _
        "y cierre todos los objetos abiertos." & vbCrLf & _
        vbCrLf & vbCrLf & _
        "Abra la ventana de depuración cuando haya terminado" & vbCrLf & _
        "para ver una lista de formularios, informes" & vbCrLf & _
        "y macros independientes" & _
        vbCrLf & vbCrLf & _
        "¿Converirt Macros?"

    If MsgBox(sMsg, vbYesNo + vbDefaultButton2 _
        , "¿Converirt Macros?") <> vbYes Then
        Exit Sub
    End If

'cerrar todos los formularios e informes abiertos
    Call CierraTodosObjetos

    iNumForms = 0
    iNumReports = 0
    iNumMacros = 0

    Access.Application.Echo False
    
    Set dicMacros = New Dictionary

        Debug.Print "Resultados búsqueda"
        Debug.Print "Tipo Objeto", "Nombre Objeto", , "Nombre Control", "Nombre Evento"
        Debug.Print String(80, "-")

        ruta = CurrentProject.Path & "\"

'Buscar las macros independientes
        For Each oScr In CurrentProject.AllMacros
            DoEvents
            Debug.Print "Macro", oScr.Name
        Next
    
'Buscar en formularios
        For Each oFrm In Application.CurrentProject.AllForms
            DoCmd.OpenForm oFrm.Name, acDesign
    
            Set frm = Forms(oFrm.Name).Form
            cForms = cForms + 1
            With frm
                For Each prp In .Properties
'Miramos las Propiedades del Formulario
                    If InStr(prp.Name, "EMMacro") > 0 Then
                        If Len(prp.Value) > 0 Then
                            Debug.Print "Form", frm.Name, , Replace(prp.Name, "EmMacro", "")
                            nomFile = ruta & frm.Name & "_" & Replace(prp.Name, "EmMacro", "") & ".scr"
                            Call Carga(nomFile, prp.Value)
                            iNumForms = iNumForms + 1
                        End If
                    End If
                Next prp
'Propiedades de los Controles del Formulario
                For Each ctl In frm.Controls
                    For Each prp In ctl.Properties
                        If InStr(prp.Name, "EMMacro") > 0 Then
                            If Len(prp.Value) > 0 Then
                                Debug.Print "Form", frm.Name, ctl.Name, Replace(prp.Name, "EmMacro", "")
                                nomFile = ruta & frm.Name & "-" & ctl.Name & "_" & Replace(prp.Name, "EmMacro", "") & ".scr"
                                Call Carga(nomFile, prp.Value)
                                iNumForms = iNumForms + 1
                            End If
                        End If
                    Next prp
                Next ctl
            End With
            DoCmd.Close acForm, oFrm.Name, acSaveNo
        Next oFrm

'Buscar en informes
        For Each oRpt In Application.CurrentProject.AllReports
            DoCmd.OpenReport oRpt.Name, acDesign
            Set rpt = Reports(oRpt.Name).Report
            cReports = cReports + 1
            With rpt
'Miramos las Propiedades del Informe
                For Each prp In .Properties
                    If InStr(prp.Name, "EmMacro") > 0 Then
                        If Len(prp.Value) > 0 Then
                            Debug.Print "Report", rpt.Name, , Replace(prp.Name, "EmMacro", "")
                            nomFile = ruta & rpt.Name & "_" & Replace(prp.Name, "EmMacro", "") & ".scr"
                            Call Carga(nomFile, prp.Value)
                            iNumReports = iNumReports + 1
                        End If
                    End If
                Next prp
'Propiedades de los Controles del Informe
                For Each ctl In rpt.Controls
                    For Each prp In ctl.Properties
                        If InStr(prp.Name, "EMMacro") > 0 Then
                            If Len(prp.Value) > 0 Then
                                Debug.Print "Report", rpt.Name, ctl.Name, Replace(prp.Name, "EmMacro", "")
                                nomFile = ruta & rpt.Name & "-" & ctl.Name & Replace(prp.Name, "EmMacro", "") & ".scr"
                                Call Carga(nomFile, prp.Value)
                                iNumReports = iNumReports + 1
                            End If
                        End If
                    Next prp
                Next ctl
            End With
            DoCmd.Close acReport, oRpt.Name, acSaveNo
        Next oRpt

        Debug.Print String(80, "-")
        Debug.Print "Búsqueda Completada"
        Debug.Print

'Iniciamos la conversión
        For Each oScr In CurrentProject.AllMacros
            DoEvents

'reconoce el cuadro de mensajes de convertir, para que no se le solicite al usuario

'Agregar manejo de errores
'Incluir comentarios

'Falso: no espera a procesar la pulsación de tecla y pasa a la siguiente sentencia

            If Not dicMacros.Exists(oScr.Name) Then
                iNumMacros = iNumMacros + 1
            End If

'seleccionamos el objeto
            DoCmd.SelectObject acMacro, oScr.Name, True
            SendKeys "{ENTER}", False
            SendKeys "{ENTER}", True
'convierte macro a vba
           
            DoCmd.RunCommand acCmdConvertMacrosToVisualBasic
    
        Next oScr

        For Each vDic In dicMacros.Keys
            DoCmd.DeleteObject acMacro, dicMacros(vDic)
        Next

        sMsg = "===*** Macros Convertidas ***===" & vbCrLf & _
            iNumMacros & " macros independientes," & vbCrLf & _
            iNumForms & " en " & cForms & " formsularios," & vbCrLf & _
            iNumReports & " en " & cReports & " informes."
    
        MsgBox sMsg, , "Conversión de Macros finalizada"
    
        Debug.Print sMsg

Error_Handler_Exit:
    On Error Resume Next
    Access.Application.Echo True
    If Not prp Is Nothing Then Set prp = Nothing
    If Not ctl Is Nothing Then Set ctl = Nothing
    If Not rpt Is Nothing Then Set rpt = Nothing
    If Not frm Is Nothing Then Set frm = Nothing
    If Not oRpt Is Nothing Then Set oRpt = Nothing
    If Not oFrm Is Nothing Then Set oFrm = Nothing
    If Not oScr Is Nothing Then Set oScr = Nothing

    If Not dicMacros Is Nothing Then Set dicMacros = Nothing

    CierraTodosObjetos
    
    Exit Sub

Error_Handler:
    Access.Application.Echo True
    
    MsgBox "Se ha producido el error" & vbCrLf & vbCrLf & _
        "Número: " & Err.Number & vbCrLf & _
        "Origen: ConversorMacros" & vbCrLf & _
        "Descripción: " & Err.Description & _
        Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Línea No: " & Erl) _
        , vbOKOnly + vbCritical, "¡Ha ocurrido un error!"
    Resume Error_Handler_Exit

End Sub

Sub Carga(nomFile As String, Valor As String)
Dim nombre As String
    
    nombre = Replace(nomFile, CurrentProject.Path & "\", "")
    nombre = Replace(nombre, ".scr", "")
    
    Open nomFile For Output As #1
    Print #1, Valor

    Close (1)
    
    Application.LoadFromText acMacro, nombre, nomFile

    dicMacros.Add nombre, nombre
    
    Kill nomFile
    
End Sub

Sub CierraTodosObjetos()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-cerrar-todos-los-formularios
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CierraTodosObjetos
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : cerrar todos los formularios a informes que estén abiertos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/api/access.allforms
'                     https://learn.microsoft.com/en-us/office/vba/api/access.allreports
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim obj As Object

    For Each obj In Application.CurrentProject.AllReports
        If obj.IsLoaded = True Then
            DoCmd.Close acReport, obj.Name, acSaveNo
        End If
    Next

    For Each obj In Application.CurrentProject.AllForms
        If obj.IsLoaded = True Then
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
    Next obj
       
    Set obj = Nothing
    
End Sub

