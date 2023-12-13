'Botón conectar de un formulario
Private Sub Conectar_Click()
Dim intret As Long
Dim atts As String
Dim strdriver As String
Dim inState As Integer
Dim contador as integer

    On Error GoTo LinError
    
'Comprueba si se ha dejado campos en blanco
    clickbutton = 0
    If IsNull(Me.DSNtxt) Or Me.DSNtxt = "" Then contador = 1: GoTo LinError
    If IsNull(Me.DBtxt) Or Me.DBtxt = "" Then contador = 2: GoTo LinError
    If IsNull(Me.Servertxt) Or Me.Servertxt = "" Then contador = 3: GoTo LinError
    If IsNull(Me.Porttxt) Or Me.Porttxt = "" Then contador = 4: GoTo LinError
    If IsNull(Me.UIDtxt) Or Me.UIDtxt = "" Then contador = 5: GoTo LinError
    If IsNull(Me.PWDtxt) Or Me.PWDtxt = "" Then contador = 6: GoTo LinError
    
    Me.visible = False
    
'Cambia el ODBC
    strdriver = "postgresql Unicode"
    atts = "DSN=" & Me.DSNtxt & Chr(0) & _
           "DATABASE=" & Me.DBtxt & Chr(0) & _
           "SERVER=" & Me.Servertxt & Chr(0) & _
           "PORT=" & Me.Porttxt & Chr(0) & _
           "UID=" & Me.UIDtxt & Chr(0) & _
           "PWD=" & Me.PWDtxt & Chr(0)
    
    intret = SQLConfigDataSource(0, 1, strdriver, atts)
    
    If FixConnections(Me.DSNtxt, Me.Servertxt, Me.DBtxt, Me.Porttxt, Me.UIDtxt, Me.PWDtxt) = False Then GoTo LinError
    
    DoCmd.Close acForm, "ParametrosConexion"
    
'Se reinicia para cargar los nuevos datos
    Restart (True)
    
    Exit Sub

LinError:
'    inState = SysCmd(acSysCmdGetObjectState, acForm, "MsgboxCritical1x1") 'Si está abierto el formulario, han fallado las tablas
    
    Select Case contador
        Case 1
            Lit1 = "Debe indicar el DSN. Rellénelo e inténtelo de nuevo" 
            Me.DSNtxt.SetFocus
        Case 2
            Lit1 = "Debe indicar el nombre de la base de datos. Rellénelo e inténtelo de nuevo" 
            Me.DBtxt.SetFocus
        Case 3
            Lit1 = "Debe indicar la IP del servidor. Rellénela e inténtelo de nuevo"
            Me.Servertxt.SetFocus
        Case 4
            Lit1 = "Debe indicar el puerto de conexión. Rellénelo e inténtelo de nuevo" 
            Me.Porttxt.SetFocus
        Case 5
            Lit1 = "Debe indicar el usuario de la base de datos. Rellénelo e inténtelo de nuevo" 
            Me.UIDtxt.SetFocus
        Case 6
            Lit1 = "Debe indicar la contraseña de accceso a la base de datos. Rellénelo e inténtelo de nuevo"
            Me.PWDtxt.SetFocus
        Case Else
            Lit1 = "Error en la conexión. Revise los parámetros e inténtelo de nuevo"
            Me.DBtxt.SetFocus
    End Select

    DoCmd.Close acForm, "ACercaDeActualizacion2"
    
    Me.visible = True
    
    If inState = 0 Then
        DoCmd.OpenForm "MsgBoxExclamation1x1"
            Form_MsgBoxExclamation1x1.msgCabecera = "¡Atención! Aviso de Mi programa" 
            Form_MsgBoxExclamation1x1.msgTxt = Lit1 & "."
            Form_MsgBoxExclamation1x1.msgTxt.FontSize = 10
            Form_MsgBoxExclamation1x1.msgOK.Caption = "Aceptar"
    Else
        clickbutton = 0
        If contador = 0 Then
            Do
                DoEvents
                If clickbutton = 1 Then
                    'Rehace la conexión que tenía
                        Call AvisoPopUp(932)
                        DoCmd.OpenForm "ParametrosConexion"
                        Exit Sub
                End If
            Loop
        End If
    End If

End Sub

'Incluir en un módulo estándar
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal _
   hwndParent As Long, ByVal fRequest As Long, ByVal _
   lpszDriver As String, ByVal lpszAttributes As String) As Long

Public Function FixConnections(DSNSource As String, ServerName As String, DatabaseName As String, port As Integer, uid As String, PWD As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-conectar-una-base-de-datos-sql/
'                     Destello formativo 391
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FixConnections
' Autor original    : Luis Viadel
' Creado            : marzo 2020
' Propósito         : establecer conexión con base de datos de PostgreSQL
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo           Descripción
'--------------------------------------------------------------------------------------------------------------------------------------------
'                     DSNSource       Obligatorio     Nombre que queremos ponerle a la conexión
'                     ServerName      Obligatorio     Nombre del servidor
'                     DatabaseName    Obligatorio     Base de datos
'                     port            Obligatorio     Puerto de conexión
'                     uid             Obligatorio     Usuario
'                     PWD             Obligatorio     Contraseña
' Retorno           : verdadero o falso dependiendo del éxito de la conexión
' Más información   : https://www.connectionstrings.com/
'                     https://learn.microsoft.com/en-GB/office/troubleshoot/access/sqlconfigdatasource-access-system-dsn
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub FixConnections_test()
' Dim hecho As Boolean
'
'    hecho = FixConnections(dsn, Server, bd, Int(port), User, Pass)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim tdfCurrent As DAO.TableDef
Dim tdfLinked As TableDef
Dim strConnectionString As String, NombreDSN As String
Dim tableOld As String, tableNew As String
Dim atts As String, strdriver As String
Dim intret As Long

    On Error GoTo LinErr
    DoCmd.SetWarnings (False)
    strConnectionString = "ODBC;DSN=" & DSNSource & ";" & _
                          "DATABASE=" & DatabaseName & ";" & _
                          "SERVER=" & ServerName & ";" & _
                          "PORT=" & port & ";" & _
                          "UID=" & uid & ";" & _
                          "PWD=" & PWD
'Creamos un formulario de evolución o progreso de la conexión    
    DoCmd.OpenForm "ACercaDeActualizacion2"
        Form_AcercaDeActualizacion2.Lite1.Caption = "Conectando con el servidor…" 
        Form_AcercaDeActualizacion2.Lite1.FontSize = 14
        Form_AcercaDeActualizacion2.Lite2.Caption = "Conectando tablas..." 
        Form_AcercaDeActualizacion2.Lite2.FontSize = 12
        Form_AcercaDeActualizacion2.Lite3.FontSize = 10
        Form_AcercaDeActualizacion2.Modal = True
    Call fPausa(0.1)
    
    For Each tdfCurrent In CurrentDb.TableDefs
        If Len(tdfCurrent.Connect) > 0 Then
            If UCase$(left$(tdfCurrent.Connect, 5)) = "ODBC;" Then
                If left(tdfCurrent.NAME, 1) = "~" Then GoTo LinNext
                If LCase(tdfCurrent.NAME) = "MiTabla" Then GoTo LinNext 'Podemos excluir las tablas que queramos en el caso de tener conexión a varias BD
                If LCase(tdfCurrent.NAME) Like "Inter_*" Then GoTo LinNext 'Podemos excluir las tablas por su nombre
                tableOld = LCase(tdfCurrent.NAME)
                TableError = tableOld
                tableNew = "public_" & LCase(tableOld)  'ELiminamos el "public" que genera postgreSQL
'Actualizamos el mensaje al usuario                
                Form_AcercaDeActualizacion2.Lite3.Caption = tableNew
                sleep(0.1)
                Set tdfLinked = CurrentDb.CreateTableDef(tableNew)
                    tdfLinked.Connect = strConnectionString
                    tdfLinked.SourceTableName = tableOld
                    CurrentDb.TableDefs.Append tdfLinked
                Set tdfLinked = Nothing
'Borra la tabla vieja y le cambia el nombre a la nueva       
                DoCmd.DeleteObject acTable, tableOld
                DoCmd.Rename tableOld, acTable, tableNew
            End If
        End If
LinNext:
    Next

    Exit Function
    
LinErr:
'Cargamos un mensaje de error personalizado    
'https://access-global.net/msgbox-personalizado/
    
    DoCmd.Close acForm, "ACercaDeActualizacion2"
    
    End 'Finaliza la ejecución
     
End Function