'Botones de comando de formulario: buscar y conectar
'Private Sub btnBuscar_Click()
Dim ruta As String

    Me.RutaUsuarios.SetFocus
    
    ruta = OpenCommDlg()
    
    If ruta <> "" Then
        Me.RutaUsuarios = ruta
        Me.RutaUsuarios.ForeColor = vbBlack
        Me.Conectar.SetFocus
    End If

End Sub

Private Sub Conectar_Click()
Dim hecho As Boolean

    If IsNull(Me.RutaUsuarios) Or Me.RutaUsuarios = "" Then Exit Sub
    
    hecho = FixConnectionsAccess(Me.RutaUsuarios)
    
    If hecho = True Then
        DoCmd.Close acForm, "Mi formulario de Conexion de Access"
    'Se reinicia para cargar los datos del nuevo usuario
        Restart (True)
    Else
        Exit Sub
    End If

End Sub

'Módulo estándar
Public Declare PtrSafe Function apiGetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OPENFILENAME As tagOPENFILENAME) As Long
Dim OPENFILENAME As tagOPENFILENAME

Public Function OpenCommDlg() As String
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-conectar-una-base-de-datos-de-access/
'                     Destello formativo 390
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : OpenCommDlg
' Autor original    : desconocido
' Creado            : desconocido
' Propósito         : manejar el cuadro de diálgo de Windows para localizar un fichero
' Argumentos        : No tiene argumentos
' Retorno           : ruta completa de la base de datos seleccionada
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub OpenCommDlg_test()
' Dim ruta As String
'
'    ruta = OpenCommDlg()
'
'    If ruta <> "" Then
'        Mensaje de error
'    End If
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Filter$, FileName$, FileTitle$, DefExt$
Dim Title$, szCurDir$

    Filter$ = "Bases de Datos (accdb)" & Chr$(0) & "Cow_Harmony_user.accdb;"
    Filter$ = Filter$ & Chr$(0)
    FileName$ = Chr$(0) & Space$(255) & Chr$(0)
    FileTitle$ = Space$(255) & Chr$(0)
    
    Title$ = "Seleccionar la Base de Usuarios ..." & Chr$(0)
    
    DefExt$ = "accdb" & Chr$(0)   'extensión por defecto
    szCurDir$ = CurDir$ & Chr$(0)  'directorio por defecto, el actual
    
    OPENFILENAME.lStructSize = Len(OPENFILENAME)
    OPENFILENAME.hwndOwner = Screen.ActiveForm.hwnd
    OPENFILENAME.lpstrFilter = Filter$
    OPENFILENAME.nFilterIndex = 1
    OPENFILENAME.lpstrFile = FileName$
    OPENFILENAME.nMaxFile = Len(FileName$)
    OPENFILENAME.lpstrFileTitle = FileTitle$
    OPENFILENAME.nMaxFileTitle = Len(FileTitle$)
    OPENFILENAME.lpstrTitle = Title$
    OPENFILENAME.flags = OFN_FILEMUSTEXIST Or OFN_READONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    OPENFILENAME.lpstrDefExt = DefExt$
    OPENFILENAME.hInstance = 0
    OPENFILENAME.lpstrCustomFilter = String(255, 0)
    OPENFILENAME.nMaxCustFilter = 255
    OPENFILENAME.lpstrInitialDir = szCurDir$
    OPENFILENAME.nFileOffset = 0
    OPENFILENAME.nFileExtension = 0
    OPENFILENAME.lCustData = 0
    OPENFILENAME.lpfnHook = 0
    OPENFILENAME.lpTemplateName = 0
    
    If apiGetOpenFileName(OPENFILENAME) <> 0 Then
        OpenCommDlg = left$(OPENFILENAME.lpstrFile, InStr(OPENFILENAME.lpstrFile, Chr$(0)) - 1)
    Else
        OpenCommDlg = ""
    End If

End Function


Public Function FixConnectionsAccess(DatabaseName As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-conectar-una-base-de-datos-de-access/
'                     Destello formativo 390
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FixConnectionsAccess
' Autor original    : Luis Viadel
' Creado            : marzo 2020
' Propósito         : establecer conexión con base de datos de Access
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo           Descripción
'--------------------------------------------------------------------------------------------------------------------------------------------
'                     DatabaseName      Obligatorio    ruta completa de la BD que queremos conectar
' Retorno           : verdadero o falso dependiendo del éxito de la conexión
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub OpenCommDlg_test()
' Dim hecho As Boolean
'
'    hecho = FixConnectionsAccess("MiRuta")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim tdfCurrent As TableDef
Dim newconn As String

    On Error GoTo ErrLinks
    
    newconn = ";Database=" & DatabaseName
    
    For Each tdfCurrent In CurrentDb.TableDefs
        If UCase$(left$(tdfCurrent.Connect, 9)) = ";DATABASE" Then
            If Len(tdfCurrent.Connect) > 0 Then
                tdfCurrent.Connect = newconn
                tdfCurrent.RefreshLink
            End If
        End If
    Next
    
    FixConnectionsAccess = True
    
    Exit Function
    
ErrLinks:
    Select Case err.Description
        Case 3011
            Lit1 = "La base de datos " & Right(newconn, Len(newconn) - 10) & " no contien la tabla " _
            & "table '" & tdfCurrent.NAME & "'. Seleccione otra base de datos."
        Case Else
            Lit1 = "Se ha producido un error: " & err.Number & "-" & err.Description
    End Select
    
    DoCmd.OpenForm "MsgBoxCritical1x1"
        Form_MsgBoxCritical1x1.Modal = True
        Form_MsgBoxCritical1x1.msgCabecera = "!ATENCIÓN! Aviso de Cow Harmony"
        Form_MsgBoxCritical1x1.msgTxt = Lit1
        Form_MsgBoxCritical1x1.msgOK.Caption = "Aceptar"
    Call DimForm(Form_MsgBoxCritical1x1)
    
    FixConnectionsAccess = False

End Function
