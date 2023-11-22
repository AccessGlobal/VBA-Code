'Módulo estándar: modImagenes
Option Compare Database
Option Explicit

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function Guardar_imagen_base64() As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-guardar-y-recuperar-imagenes/
'                     Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Guardar_imagen_base64
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 12/10/2017
' Propósito         : guardar una imagen codificada como Base64 (Incluye BBDD ejemplo)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/open-statement
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub Guardar_imagen_base64_test()
' Dim resultado As Boolean
'
'    resultado = Guardar_imagen_base64()
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fichero As String
Dim Path_inicial As String
Dim ByteImage() As Byte
Dim datos As String, ext As String
Dim rstTable As DAO.Recordset
    
    On Error GoTo LinErr
    
    Path_inicial = "C:\Cow Technologies\Access global\Destellos formativos\Destello 383\"

    fichero = mcFileDialog(Path_inicial)
    ext = Right(fichero, 3)
    
    Open fichero For Binary Access Read As #1

    ReDim ByteImage(1 To LOF(1))
        Get #1, , ByteImage
    Close #1

    datos = encodeBase64(ByteImage)

    Set rstTable = CurrentDb.OpenRecordset("imagenes")
        rstTable.AddNew
            rstTable!imagentxt = datos
            rstTable!imagenext = ext
            rstTable!imagennom = left(fichero, Len(fichero) - 4)
            rstTable!imagenfa = Format(Date, "Short date")
        rstTable.Update
    rstTable.Close
    Set rstTable = Nothing
    
    Guardar_imagen_base64 = True
    
    Exit Function
    
LinErr:
    Guardar_imagen_base64 = False
        
End Function

Public Function Leer_imagen(idimagen As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-guardar-y-recuperar-imagenes/
'                     Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Leer_imagen
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 12/10/2017
' Propósito         : recuperar una imagen codificada como Base64
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/es-es/office/client-developer/access/desktop-database-reference/stream-object-ado
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub Leer_imagen_test()
'
'    Call Leer_imagen(idimagen)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim bytimage() As Byte
Dim rstTable As DAO.Recordset
Dim nombre As String
Dim ext As String
Dim Stream

    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM imagenes WHERE idimagen=" & idimagen)
        bytimage = decodeBase64(rstTable!imagentxt)
        ext = rstTable!imagenext
    rstTable.Close
    Set rstTable = Nothing
    
    nombre = CurrentProject.Path & "\TempPicture"
    
    Set Stream = New ADODB.Stream
        Stream.Type = adTypeBinary
        Stream.Open
            Stream.Write bytimage
            Stream.SaveToFile nombre, adSaveCreateOverWrite
        Stream.Close
    Set Stream = Nothing
     
    Name nombre As nombre & "." & ext
    
    Leer_imagen = nombre & "." & ext
    
    Call ShellExecute(0&, "open", Leer_imagen, 0&, vbNullString, 1&)

End Function

Function mcFileDialog(Path_inicial As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/mcfiledialog-cuadro-de-dialogo-abrir-archivo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcFileDialog Cuadro de diálogo Abrir archivo.
' Autor original    : De varios ejemplos y artículos siendo una recopilación de ellos y de la experiencia de utilizarlo.
' Adaptado por      : Rafael Andrada | rafaelandrada@access-global.net
' Actualizado       : 10/11/2021
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objWshShell         As Object
Dim bytFileDialogType   As Byte
Dim strFileName         As String
Dim strPathFile         As String
Dim strPathInitiation   As String
Dim strRet              As String
Dim strWork             As String
      
    bytFileDialogType = 1
                                                    
    strPathInitiation = Path_inicial
                                                    
    With Application.FileDialog(bytFileDialogType)
        .Title = "Seleccionar el archivo a abrir."
    
        strPathInitiation = strPathInitiation & IIf(Right(strPathInitiation, 1) = "\", Null, "\")
        
        .InitialFileName = strPathInitiation
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Archivos PNG", "*.png*"
        .Filters.Add "JPEGs", "*.jpg"
        .Filters.Add "Bitmaps", "*.bmp"
        .FilterIndex = 3
        .AllowMultiSelect = False
        
        If Not .Show = 0 Then
            strWork = Trim(.SelectedItems.Item(1))
        End If
        
        If Not strWork = "" Then
            strFileName = Dir(strWork, vbArchive)
            strPathFile = Mid(strWork, 1, Len(strWork) - Len(strFileName))
            strRet = strPathFile & strFileName
        End If
    End With
    
    mcFileDialog = strRet
    
End Function

Public Function encodeBase64(ByRef arrData() As Byte) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-guardar-y-recuperar-imagenes/
'                     Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : encodeBase64
' Autor original    : desconocido
' Creado            : desconocido
' Propósito         : codificar una imagen como Base64
' Retorno           : nos devuelve una cadena de texto en ASCII con la codificación del fichero
' Argumento         : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo              Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     arrData()       Obligatorio       datos en bytes que se quieren codificar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub encodeBase64_test()
'Dim bytimage() As Byte
'Dim datos as string
'Dim fichero as string, ext as string
'
'    fichero = "Mi Path"
'    ext = Right(fichero, 3)
'
'    Open fichero For Binary Access Read As #1
'
'    ReDim ByteImage(1 To LOF(1))
'        Get #1, , ByteImage
'    Close #1
'
'    datos = encodeBase64(ByteImage)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objXML As MSXML2.DOMDocument60
Dim objnode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument60
        Set objnode = objXML.createElement("b64")
            objnode.DataType = "bin.base64"
            objnode.nodeTypedValue = arrData
            encodeBase64 = objnode.Text
         
        Set objnode = Nothing
    Set objXML = Nothing

End Function
 
Public Function decodeBase64(ByVal strData As String) As Byte()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-guardar-y-recuperar-imagenes/
'                     Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : decodeBase64
' Autor original    : desconocido
' Creado            : desconocido
' Propósito         : codificar una imagen como Base64
' Retorno           : decodifica una cadena codificada mediante Base64 y nos devuelve el fichero original
' Argumento         : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo              Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strData         Obligatorio       cadena codificada en ASCII que se quiere decodificar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub decodeBase64_test()
'Dim datos As String
'
'    datos = decodeBase64(Mi_campo_TXT)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objXML As MSXML2.DOMDocument60
Dim objnode As MSXML2.IXMLDOMElement
   
    Set objXML = New MSXML2.DOMDocument60
        Set objnode = objXML.createElement("b64")
            objnode.DataType = "bin.base64"
            objnode.Text = strData
            decodeBase64 = objnode.nodeTypedValue
            
        Set objnode = Nothing
    Set objXML = Nothing

End Function

Sub Guardar_imagen_Access()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-guardar-y-recuperar-imagenes/
'                     Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Guardar_imagen_Access
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 16/11/2023
' Propósito         : guardar una imagen directamente en Access en un campo datos adjuntos
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fichero As String
Dim Path_inicial As String
Dim ext As String
Dim rstTable As DAO.Recordset
Dim idimage As Integer
Dim imagenes

    Path_inicial = "C:\Cow Technologies\Access global\Destellos formativos\Destello 383\"

    fichero = mcFileDialog(Path_inicial)
    
    ext = Right(fichero, 3)
    
    Set rstTable = CurrentDb.OpenRecordset("imagenesAccess")
        rstTable.AddNew
            rstTable!imagenext = ext
            rstTable!imagennom = left(fichero, Len(fichero) - 4)
            rstTable!imagenfa = Format(Date, "Short date")
        rstTable.Update
    rstTable.Close
    Set rstTable = Nothing
    
    idimage = DMax("idimagenAccess", "imagenesAccess")
    
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM imagenesAccess WHERE idimagenAccess=" & idimage)
        rstTable.Edit
            Set imagenes = rstTable!imagen.Value
                imagenes.AddNew
                  imagenes.Fields("imagen").LoadFromFile fichero
                imagenes.Update
            Set imagenes = Nothing
        rstTable.Update
    rstTable.Close
    Set rstTable = Nothing
        
End Sub
