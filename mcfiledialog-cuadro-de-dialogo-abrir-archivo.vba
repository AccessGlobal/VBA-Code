Sub mcFileDialog_Example()

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/mcfiledialog-cuadro-de-dialogo-abrir-archivo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcFileDialog Cuadro de diálogo Abrir archivo.
' Autor original    : De varios ejemplos y artículos siendo una recopilación de ellos y de la experiencia de utilizarlo.
' Adaptado por      : Rafael Andrada .:McPegasus:. | BeeSoftware.
' Actualizado       : 10/11/2021.
' Propósito         : Mostrar el cuadro de diálogo Abrir archivo de Office para seleccionar una carpeta o un archivo.
' Retorno           : Si se convierte en función puede obtenerse diversos retornos, el nombre del archivo seleccionado o de la carpeta o incluso toda la ruta del fichero.
' Sobre Referenciar : El referenciar una librería externa nos permite seleccionar los objetos de otra aplicación que se desea que estén disponibles en nuestro código.
'                     También acceder a sus métodos utilizar las constantes.
'                     En caso de ser opcional podemos seguir utilizándolo aunque las constantes hay que sustituirlas por su valor, normalmente numérico.
'                     Más información: https://support.microsoft.com/es-es/office/add-object-libraries-to-your-visual-basic-project-ed28a713-5401-41b0-90ed-b368f9ae2513
' Referencia        : Opcional. Windows Script Host Object Model (c:\Windows\system32\wshom.ocx)/>
' Referencia        : Opcional. Microsoft Office 16.0 Object Library (c:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\...)/>
' Más información   : https://docs.microsoft.com/es-es/office/vba/api/access.application.filedialog
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar todo el procedimiento desde el Sub hasta el End Sub
'                     al portapapeles y pega en el editor de VBA de tu aplicación MS Access. Descomentar todas las líneas que nos interese (se aconseja seleccionar
'                     todas las líneas del ejemplo y utilizar el botón 'Bloque sin comentarios' de la barra de herramientas 'Edición').
'                     Pulsar F5 para ver su funcionamiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------

    '<mcKB>
    '<Grupo= Archivos/>
    '<Subgrupo=/>
    '<Autor= Rafael .:McPegasus:./>
    '<Contacto= rafael@mcpegasus.net, www.mcpegasus.net/>
    '
    'OBSERVACIONES INTERNAS PARA DEVELOPER (visibles sólo aquí): Sólo se guarda en tblmcKB este primer grupo donde se indica la versión <Versión= 01/>, el grupo de <Revisión = nn/> no.
    '<Versión= 01/><Fecha creación= 10/11/2021/>
    '<Descripción= Ejemplo de abrir un cuadro de diálogo para seleccionar una carpeta o un archivo./>
    '
    '<Tags= abrir archivo, guardar archivo, seleccionar carpeta./>
    '
    '<Referencia= Opcional: Windows Script Host Object Model (c:\Windows\system32\wshom.ocx)/>
    '<Referencia= Opcional: Microsoft Office nn.0 Object Library (c:\Program Files (x86)\Common Files\Microsoft Shared\OFFICEnn\...)/>
    '
    '<Más información= https://docs.microsoft.com/es-es/office/vba/api/access.application.filedialog/>
    '
    '</mcKB>
    
    Dim objWshShell                                 As Object           'En mcstrSpecialFolderPath hay más información sobre esta declaración.
'    Dim objWshShell                                 As New WshShell

    Dim bytFileDialogType                           As Byte

    Dim strFileName                                 As String
    Dim strPathFile                                 As String
    Dim strPathInitiation                           As String
    Dim strRet                                      As String           'Valor del retorno final.
    Dim strWork                                     As String

    
'    bytFileDialogType = 1
    'msoFileDialogOpen = 1 = Cuadro de diálogo Abrir. _
        Nos permite hacer una selección de un archivo. Nos retorna la ruta completa del archivo. _
        Aparentemente hace lo mismo que el 3.
    
'    bytFileDialogType = 2
    'msoFileDialogSaveAs = 2 = Cuadro de diálogo Guardar como. _
        La típica selección "Guardar como" que nos sirve para guardar un documento que hayamos generado, un pdf, un Excel por ejemplo. _
        Si se indica el nombre y no existe en la carpeta que se elija finaliza la acción, pero si se selecciona un archivo que ya exista nos lo indica y nos pregunta ¿Desea reemplazarlo?
    
'    bytFileDialogType = 3
    'msoFileDialogFilePicker = 3 = Cuadro de diálogo selector de archivos. _
        Aparentemente hace lo mismo que el 3.
    
    bytFileDialogType = 3
    'msoFileDialogFolderPicker = 4 = Cuadro de diálogo Selector de carpetas.    Si tenemos la referencia Microsoft Office nn.0 Object Library se puede utilizar la constante msoFileDialogFolderPicker.
                                                    
    strPathInitiation = ""                                                      'Si no se pasa una ruta inicial se abre en la ruta \Mis documentos.
                                                    
    With Application.FileDialog(bytFileDialogType)
        Select Case bytFileDialogType                                           'Titulo de la barra de títulos.
            Case 1
                .Title = "Seleccionar el archivo a abrir."
        
            Case 2
                .Title = "Seleccionar la carpeta o archivo donde guardar como."
            
            Case 3
                .Title = "Seleccionar el archivo a importar."
        
            Case 4
                .Title = "Seleccionar la carpeta donde alojar el pdf."
        
        End Select
        
        If strPathInitiation = "" Then
            Set objWshShell = CreateObject("WScript.Shell")
            strPathInitiation = objWshShell.SpecialFolders("MyDocuments")
            Set objWshShell = Nothing

        End If
        strPathInitiation = strPathInitiation & IIf(Right(strPathInitiation, 1) = "\", Null, "\")
        
        .InitialFileName = strPathInitiation
        If bytFileDialogType = 1 Or bytFileDialogType = 3 Then
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
            .Filters.Add "Archivos de Excel", "*.xl*"
            .Filters.Add "Archivos de Word", "*.doc*"
            .Filters.Add "JPEGs", "*.jpg"
            .Filters.Add "Bitmaps", "*.bmp"
            .FilterIndex = 3
            .AllowMultiSelect = False
        
        End If
        
        '28/11/2021 No funciona .InitialView en Windows 10. Si se sabe como hacerlo funcionar correctamente ruego se ponga en contacto en rafael@mcpegasus.net.
        'Establecer el tipo de Vistas, nombre de las constantes y su valor decimal para usar cuando no se referencia. _
        msoFileDialogViewLargeIcons = 6, Iconos grandes. _
        msoFileDialogViewSmallIcons = 7, Iconos pequeños _
        msoFileDialogViewList = 1, Lista _
        msoFileDialogViewDetails = 2, Detalles _
        msoFileDialogViewProperties = 3, Propiedades _
        msoFileDialogViewPreview = 4, Vista previa _
        msoFileDialogViewThumbnail = 5, Vista en miniaturas. _
        msoFileDialogViewWebView = 8, Vista web _
        msoFileDialogViewTiles = 9, Se produce un error en Windows XP Pro SP2, Office XP.
        '.InitialView = 9
        
        If Not .Show = 0 Then
            strWork = Trim(.SelectedItems.Item(1))
        
        End If
        
        If Not strWork = "" Then
            strFileName = Dir(strWork, vbArchive)
            strPathFile = Mid(strWork, 1, Len(strWork) - Len(strFileName))
            
            If bytFileDialogType = 1 Or bytFileDialogType = 2 Or bytFileDialogType = 3 Then
                strRet = strPathFile & strFileName
                
            Else
                strRet = strPathFile
                
            End If
        End If
    End With
    
    Debug.Print strRet
    
End Sub
