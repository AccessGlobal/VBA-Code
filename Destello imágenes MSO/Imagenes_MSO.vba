'Módulo estándar: modgraficos
Option Compare Database
Option Explicit

Private Const MSV_NOMBRE_MODULO As String = "modGraficos"

Public Sub PutCmdMsoImage(obCmd As CommandButton, pMso As String, Optional pSize As Long = 32)
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-imagenes-mso/
'                     Destello formativo 395
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : PutCmdMsoImage
' Autor original    : Alba Salvá | albasalvaaccess-global.net
' Creado            : 20/05/2020
' Propósito         : colocar la imagen seleccionda en el botón
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     obCmd         Obligatorio      objeto sobre el que ponemos la imagen
'                     pMso          Obligatorio      nombre de la imagen
'                     obCmd         pSize            tamaño
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub PutCmdMsoImage_test()
'
'      PutCmdMsoImage MiBotón, imgCmd, 16
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim lImage As Object
    
    On Error GoTo lbError
    
    If GetPictureFromMso(pMso, lImage, pSize) Then
        On Error Resume Next
        If Dir(CurrentProject.Path & "\small.bmp") <> "" Then
            Kill CurrentProject.Path & "\small.bmp"
        End If
        On Error GoTo lbError
        
'        Debug.Print lImage.Type
        SavePicture lImage, CurrentProject.Path & "\small.bmp"
        obCmd.Picture = CurrentProject.Path & "\small.bmp"
        
    End If
    Set lImage = Nothing
    
    GoTo lbFinally

lbError:
    MsgBox Err & vbCrLf & Err.Description
    Stop
    Resume
    
lbFinally:
'    Kill CurrentProject.Path & "\small.bmp"
        
    On Error GoTo 0
    
End Sub


Public Function GetPictureFromMso(pMso As String, PImage As Object, Optional pSize As Long = 32) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-imagenes-mso/
'                     Destello formativo 395
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetPictureFromMso
' Autor original    : Alba Salvá | albasalvaaccess-global.net
' Creado            : 20/05/2020
' Propósito         : crear o modificar un origen de datos remoto
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     obCmd         Obligatorio      objeto sobre el que ponemos la imagen
'                     pMso          Obligatorio      nombre de la imagen
'                     obCmd         pSize            tamaño
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub PutCmdMsoImage_test()
'
'      PutCmdMsoImage MiBotón, imgCmd, 16
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    
Dim C As Office.CommandBars
Dim o As Object
' Creación del menú para los elementos
    On Error GoTo ErrorTrap
    
    #If Access2000 = False Then
        Set C = CurrentProject.Application.CommandBars
        Set o = C.GetImageMso(pMso, pSize, pSize)
    #End If
    
    If Not o Is Nothing Then
        Set PImage = o
        GetPictureFromMso = True
    Else
        GetPictureFromMso = False
        Set PImage = Nothing
    End If
    
    Exit Function

ErrorTrap:
    GetPictureFromMso = False
    Set PImage = Nothing

End Function

'Formulario
Public Sub AsignaImagenBoton()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-imagenes-mso/
'                     Destello formativo 395
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AsignaImagenBoton
' Autor original    : Alba Salvá | albasalvaaccess-global.net
' Creado            : 20/05/2020
' Propósito         : asignar una imagen mso a un botón
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim imgCmd As String
    
    DoCmd.Hourglass True
    imgCmd = "Font"
    
    If Trim(Me.txtFichero & "") <> "" Then
        Select Case LCase(Trim(Mid(Me.txtFichero, InStrRev(Me.txtFichero, ".") + 1)))
            Case "mdb" ', "accdb"
                imgCmd = "FileSaveAsAccess2000" ' "MicrosoftAccess"
                
            Case "accdb"
                imgCmd = "FileSaveAsAccess2007"
                
            Case "mda", "accda"
                imgCmd = "AddInsMenu"
            
            Case "accdr"
                imgCmd = "DatabaseMakeMdeFile"
            
            Case "xls" ', "xlsx", "xlsm"
                imgCmd = "FileSaveAsExcel97_2003" '"MicrosoftExcel"
            
            Case "xlsx"
                imgCmd = "FileSaveAsExcelXlsx"
            
            Case "xlsm"
                imgCmd = "FileSaveAsExcelXlsxMacro"
            
            Case "xlsb"
                imgCmd = "FileSaveAsExcelXlsb"
                
            Case "doc" ', "docx", "docm"
                imgCmd = "FileSaveAsWord97_2003" '"MailMergeClearMergeType"
            
            Case "docx", "docm"
                imgCmd = "FileSaveAsWordDocx"
            
            Case "dot", "dotx", "dotm"
                imgCmd = "FileSaveAsWordDotx"
                
            Case "vbs", "cmd", "bat" ', "ps1"
                imgCmd = "AddInManager"
            
            Case "ps1"
                imgCmd = "MacroRun"
            
            Case "ppt" ', "pptx"
                imgCmd = "FileSaveAsPowerPoint97_2003" '"MicrosoftPowerPoint"
                
            Case "pptx", "ppsx"
                imgCmd = "FileSaveAsPowerPointPptx"
            
            Case "sql"
                imgCmd = "AdpViewSqlPane" '"_3DModelSceneGallery"
                
            Case "txt"
                imgCmd = "TextFromFileInsert"
            
            Case Else
                imgCmd = "Help"
        
        End Select
    
    End If
    
    PutCmdMsoImage Me.cmdTest, imgCmd, 16
    
    On Error Resume Next
    Kill CurrentProject.Path & "\small.bmp"
    On Error GoTo 0
    DoCmd.Hourglass False
    
End Sub

Private Sub cboSize_Click()

    Me.img16.Visible = Me.cboSize >= 16
    Me.img32.Visible = Me.cboSize >= 32

    Me.img48.Visible = Me.cboSize >= 48
    Me.img64.Visible = Me.cboSize >= 64

    Me.img96.Visible = Me.cboSize >= 96
    Me.img128.Visible = Me.cboSize >= 128
    
    Me.cmd16.Visible = Me.cboSize >= 16
    Me.cmd32.Visible = Me.cboSize >= 32

    Me.cmd48.Visible = Me.cboSize >= 48
    Me.cmd64.Visible = Me.cboSize >= 64
    
    Me.cmdTest16.Visible = Me.cboSize >= 16
    Me.cmdTest32.Visible = Me.cboSize >= 32
    Me.cmdTest48.Visible = Me.cboSize >= 48
    Me.cmdTest64.Visible = Me.cboSize >= 64

End Sub

Private Sub cboVersion_Click()

    Me.lstImages.RowSource = "SELECT imageMso FROM ImageMsoVersiones WHERE " & Me.cboVersion & " = True ORDER BY imageMso"
    
End Sub

Private Sub cmdBuscar_Click()
Dim BF As Object
Const msoFileDialogFilePicker = 3
    
'Seleccionamos un fichero
    Set BF = FileDialog(msoFileDialogFilePicker)
    
        With BF
            .AllowMultiSelect = False
            .ButtonName = "Seleccionar"
            .InitialFileName = IIf(Me.txtFichero & vbNullString = vbNullString, "C:\", Me.txtFichero)
            .Title = "Buscar fichero"
            If .Show = -1 Then
                Me.txtFichero = .SelectedItems(1)
            End If
        End With
    
    Set BF = Nothing
    
    AsignaImagenBoton
    
End Sub

Private Sub Form_Load()
    
    Me.cboVersion = "O_2010"
    cboVersion_Click
    Me.cboSize = 16
    cboSize_Click
    
End Sub

Private Sub lstImages_Click()

    Set Me.img16.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 16, 16)
    Set Me.img32.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 32, 32)

    Set Me.img48.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 48, 48)
    Set Me.img64.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 64, 64)

    Set Me.img96.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 96, 96)
    Set Me.img128.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 128, 128)

    PutCmdMsoImage Me.cmdTest16, Me.lstImages, 16 '"AddInsMenu", 16
    PutCmdMsoImage Me.cmdTest32, Me.lstImages, 32 '"FileSaveAsExcelXlsxMacro", 32
    PutCmdMsoImage Me.cmdTest48, Me.lstImages, 48 '"AddInManager", 48
    PutCmdMsoImage Me.cmdTest64, Me.lstImages, 64
    
    Set Me.cmd16.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 16, 16)
    Set Me.cmd32.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 32, 32)

    Set Me.cmd48.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 48, 48)
    Set Me.cmd64.Picture = Application.CommandBars.GetImageMso(Me.lstImages, 64, 64)

End Sub

Private Sub txtFichero_AfterUpdate()

    AsignaImagenBoton

End Sub

