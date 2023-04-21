'Código del botón
Private Sub btnCrear_Click()
Dim App As New Access.Application
Dim strPathNameNew As String
Dim strPathNameNewe As String
Dim strPathdbs As String


    strPathdbs = Application.CurrentProject.Path & "\" & Application.CurrentProject.Name
    strPathNameNew = Application.CurrentProject.Path & "\MiApp_tmp.accdb"
        
    If Not Dir(strPathNameNew, vbArchive) = "" Then
        Kill strPathNameNew
    End If
        
    strPathNameNewe = Application.CurrentProject.Path & "\MiApp.accde"
        
    If Not Dir(strPathNameNewe, vbArchive) = "" Then
        Kill strPathNameNewe
    End If
        
    Call kbCopyFile(strPathdbs, strPathNameNew)
    
'Configurar las Propiedades de Inicio del programa
    Call mcBeginProperties(strPathNameNew)
        
    App.AutomationSecurity = 1 'msoAutomationSecurityLow
    
    App.SysCmd 603, strPathNameNew, strPathNameNewe
    
    Set App = Nothing
    
    Kill strPathNameNew
    
'Introduce tu mensaje personalizado
    MsgBox "La nueva versión ha sido creada con éxito"

End Sub




'Código a incorporar en un módulo estándar
Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-accde-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Módulo original   : mcCreateNewAccde
' Autor             : McPegasus|www.mcpegasus.net|rafael@mcpegasus.net
' Creado            : 07/03/2007
' Revisión          : 05/09/2019, 24/07/2019
' Propósito         : crear un fichero accde con las características que queramos
' Funciones         : mcCreateaccdePackage
'                     mcCopiar
'                     kbCopyFile
'                     mcPropertiesdbs
'                     mcBeginProperties
'                     mcBeginPropertiesII
'-----------------------------------------------------------------------------------------------------------------------------------------------
Const cstrAppTitle     As String = "Nombre de mi App"
Const cstrTitle        As String = "Título"
Const cstrProjectName  As String = "Nombre del proyecto VBA"
Const cstrAuthor       As String = "Autor"
Const cstrCompany      As String = "Empresa"
Const cstrEmail        As String = "Mi email"
Const cstrWeb          As String = "Mi Web"
Const cstrPoblación    As String = "mi población"
Const cstrComments     As String = "Mis comentarios"
Const cstrStartupForm  As String = "frmLogin"
Const cstartupRibbon   As String = "Ribbon que carga inicialmente"

Public Sub mcCreateaccdePackage()
   
    Call mcCopiar

'Aquí puedes añadir código que necesites en la nueva versión, como por ejemplo, borrar
'tablas temporales, ...

End Sub

Sub mcCopiar()
Dim App As New Access.Application
Dim strPathNameNew As String
Dim strPathNameNewe As String
Dim strPathdbs As String
    
    strPathdbs = Application.CurrentProject.Path & "\" & Application.CurrentProject.Name
    strPathNameNew = Application.CurrentProject.Path & "\MiApp_tmp.accdb"
        
    If Not Dir(strPathNameNew, vbArchive) = "" Then
        Kill strPathNameNew
    End If
        
    strPathNameNewe = Application.CurrentProject.Path & "\MiApp.accde"
        
    If Not Dir(strPathNameNewe, vbArchive) = "" Then
        Kill strPathNameNewe
    End If
        
    Call kbCopyFile(strPathdbs, strPathNameNew)
    
'Configurar las Propiedades de Inicio del programa
    Call mcBeginProperties(strPathNameNew)
        
    App.AutomationSecurity = 1 'msoAutomationSecurityLow
    
    App.SysCmd 603, strPathNameNew, strPathNameNewe
    
    Set App = Nothing
    
    Kill strPathNameNew
        
    MsgBox "La nueva versión ha sido creada con éxito"

End Sub

Function kbCopyFile(ByVal Source$, ByVal Destination$) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-accde-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : kbCopyFile
' Autor original    : desconocido
' Creado            : desconocido
' Propósito         : facilitar la copia de un fichero
' Retorno           : valor long que indica el tamaño del fichero copiado
' Argumento/s       : La sintaxis de la función consta de los siguientes argumentos:
'                     Parte                   Modo                    Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Source$                 Obligatorio             Path origen donde se encuentra el fchero a copiar
'                     Destination$            Obligatorio             Path destino donde se grabará el fchero copiado
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : http://support.microsoft.com/kb/102671/es
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim Index1 As Integer, NumBlocks As Integer
Dim SourceFile As Integer, DestFile As Integer
Dim FileLength As Long, LeftOver As Long, AmountCopied As Long
Dim FileData As String
Const BlockSize = 32768
    

    On Error GoTo Err_kbCopyFile

' Remove the destination file.
    DestFile = FreeFile
    
    Open Destination For Output As DestFile
    
    Close DestFile

' Open the source file to read from.
    SourceFile = FreeFile
    Open Source For Binary Access Read As FreeFile

' Open the destination file to write to.
    DestFile = FreeFile
    Open Destination For Binary As DestFile

' Get the length of the source file.
    FileLength = LOF(SourceFile)

' Calculate the number of blocks in the file and left over.
    NumBlocks = FileLength \ BlockSize
    LeftOver = FileLength Mod BlockSize

' Create a buffer for the leftover amount.
    FileData = String$(LeftOver, 32)

' Read and write the leftover amount.
    Get SourceFile, , FileData
    Put DestFile, , FileData

' Create a buffer for a block to be read.
    FileData = String$(BlockSize, 32)

' Read and write the remaining blocks of data.
    For Index1 = 1 To NumBlocks
' Read and write one block of data.
        Get SourceFile, , FileData
        Put DestFile, , FileData
    Next Index1

    Close SourceFile, DestFile
    kbCopyFile = AmountCopied
    
Bye_kbCopyFile:
        Exit Function
    
Err_kbCopyFile:
        kbCopyFile = -1 * err
        Resume Bye_kbCopyFile

End Function

Public Sub mcPropertiesdbs()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-accde-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcPropertiesdbs
' Autor             : McPegasus|www.mcpegasus.net|rafael@mcpegasus.net
' Creado            : 07/03/2007
' Revisión          : 05/09/2019, 24/07/2019
' Propósito         : cambiar las propiedades de la nueva base de datos
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim strNameProperty As String
Dim strValueProperty As String
    
    On Error GoTo Err_CapturarError

'Título de la aplicación.
    strNameProperty = "AppTitle"
    strValueProperty = cstrAppTitle
    
    On Error Resume Next
    CodeDb.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Properties.Append CodeDb.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Nombre de Proyecto.
'En caso de ser el mismo valor se produce el error 3709 La clave de búsqueda no se encontró en ningún registro.

    strNameProperty = cstrProjectName
    
    On Error Resume Next
        
    Application.SetOption "Project Name", strNameProperty
    On Error GoTo 0
    
'Título de la Base de Datos.
    strNameProperty = "Title"
    strValueProperty = cstrTitle
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Autor.
    strNameProperty = "Author"
    strValueProperty = cstrAuthor
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Organización.
    strNameProperty = "Company"
    strValueProperty = cstrCompany
    
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
        
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Último Cambio.
    strNameProperty = "Category"
    strValueProperty = Format(Date, "mmm yyyy")
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Número Versión.
    strNameProperty = "Keywords"
    strValueProperty = version
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Web.
    strNameProperty = "pbdpWeb"
    strValueProperty = cstrWeb
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Delete (strNameProperty)
        
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Append _
            CodeDb.Containers!Databases.Documents!UserDefined.CreateProperty(strNameProperty, dbText, strValueProperty)

'Email.
    strNameProperty = "pbdpEmail"
    strValueProperty = cstrEmail
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Append _
            CodeDb.Containers!Databases.Documents!UserDefined.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Población.
    strNameProperty = "pbdpPoblación"
    strValueProperty = cstrPoblación
    
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    CodeDb.Containers!Databases.Documents!UserDefined.Properties.Append _
            CodeDb.Containers!Databases.Documents!UserDefined.CreateProperty(strNameProperty, dbText, strValueProperty)
    
'Comentario.
    strNameProperty = "Comments"
    strValueProperty = cstrComments
        
    On Error Resume Next
    CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Delete (strNameProperty)
    
    On Error GoTo 0
    
    If Not strValueProperty = "" Then
        CodeDb.Containers!Databases.Documents!SummaryInfo.Properties.Append _
            CodeDb.Containers!Databases.Documents!SummaryInfo.CreateProperty(strNameProperty, dbText, strValueProperty)
    End If
    
Salida:
        Exit Sub
    
Err_CapturarError:
    Select Case err.Number
        Case Else
            'Cazar todos aquellos errores inesperados.
            MsgBox err.Number & " " & err.Description, vbCritical, "En mcPropertiesdbs."
        End Select
    Resume Salida           'Salida a otro procedimiento.

End Sub

Public Sub mcBeginProperties(ByVal strPathNamedbs As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-accde-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcBeginProperties y mcBeginPropertiesII
' Autor             : McPegasus|www.mcpegasus.net|rafael@mcpegasus.net
' Creado            : 07/03/2007
' Revisión          : 05/09/2019, 24/07/2019
' Propósito         : seleccionar la propiedad a cambiar y llamar a la función mcBeginPropertiesII para realizar el cambio
' Argumento         : La sintaxis de la función consta del siguiente argumento:
'                     Parte                   Modo                    Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strPathNamedbs          Obligatorio             Path de la base de datos que queremos modificar
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim strNameProperty As String
  
    On Error GoTo Err_CapturarError
    
'Mostrar fichas de documentos.
    strNameProperty = "ShowDocumentTabs"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, True)
        
'Permitir el uso de menús no restringidos.
    strNameProperty = "AllowFullMenus"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Permitir el uso de menús contextuales predeterminados.
    strNameProperty = "AllowShortcutMenus"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Permitir mostrar código después de un error
    strNameProperty = "AllowBreakIntoCode"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Mostrar el formulario de inicio
    strNameProperty = "StartupForm"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, True)
    
'Mostrar la banda de opciones
    strNameProperty = "CustomRibbonID"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, True)
    
'Presentar la ventana Base De Datos al Iniciar.
    strNameProperty = "StartUpShowDBWindow"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Presentar la Barra de Estado.
    strNameProperty = "StartUpShowStatusBar"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Permitir el uso de las barras de herramientas incorporadas.
    strNameProperty = "AllowBuiltInToolbars"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Permitir cambios en barras de herramientas y menús.
    strNameProperty = "AllowToolbarChanges"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
    
'Usar las teclas especiales de Access.
    strNameProperty = "AllowSpecialKeys"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)
        
'Permitir Ignorar Inicio (Activar la tecla shift)
    strNameProperty = "AllowBypassKey"
    Call mcBeginPropertiesII(strPathNamedbs, strNameProperty, False)

Salida:
    Exit Sub

Err_CapturarError:
    Select Case err.Number
        Case Else
            'Cazar todos aquellos errores inesperados.
            MsgBox err.Number & " " & err.Description, vbCritical, "En mcBeginProperties."
    
    End Select
    Resume Salida                                           'Salida a otro procedimiento.

End Sub

Private Sub mcBeginPropertiesII(ByVal strPathNamedbs As String, ByVal strNameProperty As String, ByVal blnTrue As Boolean)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-accde-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcBeginProperties y mcBeginPropertiesII
' Autor             : McPegasus|www.mcpegasus.net|rafael@mcpegasus.net
' Creado            : 07/03/2007
' Revisión          : 05/09/2019, 24/07/2019
' Propósito         : seleccionar la propiedad a cambiar y llamar a la función mcBeginPropertiesII para realizar el cambio
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos:
'                     Parte                   Modo                    Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strPathNamedbs          Obligatorio             Path de la base de datos que queremos modificar
'                     strNameProperty         Obligatorio             Nombre de la propiedad que deseamos cambiar
'                     blnTrue                 Obligatorio             booleano que indica el valor verdadero/false de la propiedad
'---------------------------------------------------------------------------------------------------------------------------------------------------

Dim dbs As DAO.Database

    On Error GoTo Err_CapturarError
    
    Set dbs = OpenDatabase(strPathNamedbs, True, False)
        
    On Error Resume Next
    
    dbs.Properties(strNameProperty) = blnTrue
    
    On Error GoTo 0
    
    If strNameProperty = "StartupForm" Then
        dbs.Properties.Append dbs.CreateProperty(strNameProperty, dbText, cstrStartupForm)
    End If
    
    If strNameProperty = "CustomRibbonID" Then
       dbs.Properties.Append dbs.CreateProperty(strNameProperty, dbText, cstartupRibbon)
    End If

Salida:
    Exit Sub

Err_CapturarError:
    Select Case err.Number
        Case 3356 'Intenta abrir una dbs y está ocupada, intentar de nuevo después un breve tiempo.
            'Utlizar la función sleep, por ejemplo
                
        Case Else
            'Cazar errores inesperados.
            MsgBox err.Number & " " & err.Description, vbCritical, "En mcBeginPropertiesII."
            
    End Select
    
End Sub
