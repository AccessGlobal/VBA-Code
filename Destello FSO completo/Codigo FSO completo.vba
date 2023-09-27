Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-objeto-fso-recopilacion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mod_FSO
' Autor original    : Alba Salvá
' Creado            : diferentes fechas
' Propósito         : mostrar el uso de todas las posibilidad del objeto FSO en un único ejemplo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
' Más información   : diferentes funciones que recorren el PC del usuario y van creando un fichero de texto con todo lo que este contiene.
'                     - Directorios
'                     - Carpetas
'                     - Ficheros y sus propiedades
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim ts As Object
Dim fl As Object

Dim MiForm As Form_Principal
Dim SumaBytes As Currency

Const ForWriting = 2
Const ForAppending = 8
Const TristateUseDefault = &HFFFFFFFE '-2

Sub RecorrePC()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : RecorrePC
' Autor original    : Alba Salvá
' Creado            : diferentes fechas
' Propósito         : recorrer el PC del usuario extrayendo unidades y sus propiedades
' Retorno           : va creando un archivo txt con lo que va encontrando
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723602(v=vs.85)?redirectedfrom=MSDN
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub recorrerpc_test()
'
'        Call recorrerPC
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Discos As Object
Dim Disco As Object
Dim strMsg As String
    
    Set MiForm = Form_Principal
    
    MiForm.BarMin = 0
    MiForm.BarValue = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists("C:\Listado_PC") Then fso.CreateFolder "C:\Listado_PC"
    fso.CreateTextFile "C:\Listado_PC\Listado_PC.txt"
    Set fl = fso.GetFile("C:\Listado_PC\Listado_PC.txt")
    
    Set ts = fl.OpenAsTextStream(ForWriting, TristateUseDefault)
    ts.WriteLine "Listado de unidades, carpetas y archivos, y sus características."
    ts.WriteLine "================================================================"
    ts.WriteBlankLines 3
    
    Set Discos = fso.Drives
    
    ts.WriteLine "Tu PC tiene " & Discos.Count & " unidades de disco"
    ts.Close
    
    For Each Disco In Discos
        DoEvents
          
        strMsg = "C:\Listado_PC\Unidad " & Disco.DriveLetter & ".txt"
        MiForm.txtUnidad = Disco.Path
        MiForm.txtFichero = strMsg
        
        fso.CreateTextFile strMsg
        Set fl = fso.GetFile(strMsg)
    
        Set ts = fl.OpenAsTextStream(ForAppending, TristateUseDefault)
        
        ts.WriteBlankLines 1
        
        With Disco
                        
            ts.Write vbTab & .DriveLetter & " - "
            Select Case .DriveType
                Case 1
                    ts.Write "Removible"
                Case 2
                    ts.Write "Fijo"
                Case 3
                    ts.Write " en Red"
                Case 4
                    ts.Write "CDRom"
                Case 5
                    ts.Write "Disco RAM"
                Case Else
                    ts.Write "Desconocido"
            End Select
            
            If .DriveType = 3 Then
                
                ts.WriteLine "Recurso de red: " & .ShareName
            Else
                If .IsReady Then
                    ts.WriteLine vbTab & "Nombre: " & .VolumeName
                Else
                    ts.WriteLine vbTab & "Unidad no disponible"
                End If
            End If
            
            ts.WriteLine vbTab & vbTab & "Está activo: " & .IsReady
            If .IsReady Then
                ts.WriteLine vbTab & vbTab & "Nº de serie: " & .SerialNumber
                ts.WriteLine vbTab & vbTab & "Sistema de srchivos: " & .FileSystem
                ts.WriteLine vbTab & vbTab & "Capacidad total: " & Format(.TotalSize, "#,##0") & " bytes"
                ts.WriteLine vbTab & vbTab & "Espacio libre  : " & Format(.FreeSpace, "#,##0") & " bytes"
                ts.WriteLine vbTab & vbTab & "Carpeta raiz: " & .RootFolder
                
                MiForm.BarMax = .TotalSize - .FreeSpace
            
            End If
            ts.WriteLine vbTab & vbTab & "Ruta: " & .Path
            ts.Close
            
            If .IsReady Then recorreCarpetas .RootFolder
            
        End With
'Para listar sólo la primera unidad, quita el comentario
'Stop
        
    Next

    Set fso = Nothing
    
End Sub

Sub RecorreCarpetas(strCarpeta As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : RecorreCarpetas
' Autor original    : Alba Salvá
' Creado            : diferentes fechas
' Propósito         : recorrer las diferentes carpetas del PC del usuario extrayendo propiedades de las mismas, obteniendo así, el árbol de
'                     carpetas
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strCarpeta      Obligatorio      Carpeta PADRE desde la que queremos extrar su árbol
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : va creando un archivo txt con lo que va encontrando
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub recorreCarpetas_test( unidadPadre)
'
'        Call recorrerPC
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    Static nivel As Integer, n As Integer
    Dim subCarpeta As Object, Fichero As Object
    Dim MiFile As Object
    Dim Saltos As String
    
    On Error Resume Next
    
    Set ts = fl.OpenAsTextStream(ForAppending, TristateUseDefault)
    
    MiForm.txtRuta = strCarpeta
    
    Set subCarpeta = fso.GetFolder(strCarpeta)
    
    MiForm.txtCFiles = subCarpeta.Files.Count
    MiForm.txtcSize = Format(subCarpeta.Size, "#,##0") & " bytes"
    MiForm.txtCFolders = subCarpeta.SubFolders.Count
    
    If subCarpeta.Name <> "Listado_PC" Then
        For n = 0 To nivel + 1
            Saltos = Saltos & vbTab
        Next
        
        ts.WriteBlankLines 1
        ts.WriteLine "======= DATOS CARPETA ========"
        
        With subCarpeta
            
            Set ts = fl.OpenAsTextStream(ForAppending, TristateUseDefault)
            
            ts.WriteBlankLines 1
            
            ts.WriteLine Saltos & "Nombre           : " & .Name
            ts.WriteLine Saltos & "Nombre corto     : " & .ShortName
            ts.WriteLine Saltos & "Ruta             : " & .Path
            ts.WriteLine Saltos & "Ruta corta       : " & .ShortPath
            ts.WriteLine Saltos & "Tipo             : " & .Type
            ts.WriteLine Saltos & "Atributos        : " & sacaAtributos(.Attributes)
            ts.WriteLine Saltos & "Creada el        : " & .DateCreated
            ts.WriteLine Saltos & "Modificada el    : " & .DateLastModified
            ts.WriteLine Saltos & "Último acceso el : " & .DateLastAccessed
            
            ts.WriteLine Saltos & "Carpeta superior : " & .ParentFolder
            
            ts.WriteLine Saltos & "Coniene          : " & subCarpeta.Files.Count & " fichero(s) y " & .SubFolders.Count & " subcarpeta(s)"
            ts.WriteLine Saltos & "Utiliza          : " & Format(.Size, "#,##0") & " bytes"
            
        End With
    End If
    
    ts.WriteBlankLines 1
    ts.WriteLine "=============================="
    
    If subCarpeta.Files.Count > 0 Then
        ts.WriteBlankLines 2
        ts.WriteLine "========== FICHEROS =========="
    End If
    
    
    On Error GoTo 0
    
    Echo False
    If subCarpeta.Files.Count > 0 Then
        Saltos = ""
        nivel = nivel + 1
        For n = 0 To nivel + 1
            Saltos = Saltos & vbTab
        Next

        For Each Fichero In subCarpeta.Files
            DoEvents

            Set MiFile = fso.GetFile(Fichero)
            On Error Resume Next
            With MiFile
                
                ts.WriteBlankLines 1
                ts.WriteLine Saltos & "Nombre          : " & .Name
                ts.WriteLine Saltos & "Nombre corto    : " & .ShortName
                ts.WriteLine Saltos & "Ruta            : " & .Path
                ts.WriteLine Saltos & "Ruta corta      : " & .ShortPath
                ts.WriteLine Saltos & "Nombre base     : " & fso.GetBaseName(.Path)
                ts.WriteLine Saltos & "Extnsión        : " & fso.GetExtensionName(.Path)
                ts.WriteLine Saltos & "Tipo            : " & .Type
                ts.WriteLine Saltos & "Atributos       : " & sacaAtributos(.Attributes)
                ts.WriteLine Saltos & "Carpeta         : " & .ParentFolder
                ts.WriteLine Saltos & "Creado el       : " & .DateCreated
                ts.WriteLine Saltos & "Modificado el   : " & .DateLastModified
                ts.WriteLine Saltos & "Último acceso el: " & .DateLastAccessed
                ts.WriteLine Saltos & "Tamaño          : " & Format(.Size, "#,##0") & " bytes"
                   
                SumaBytes = SumaBytes + .Size
                MiForm.BarValue = SumaBytes
                
            End With
        Next
        nivel = nivel - 1
    End If
    Echo True
    
    Set subCarpeta = Nothing
    
    If fso.GetFolder(strCarpeta).Files.Count > 0 Then
        ts.WriteBlankLines 1
        ts.WriteLine "=============================="
    End If
    
    If fso.GetFolder(strCarpeta).SubFolders.Count > 0 Then
        
        ts.WriteBlankLines 2
        ts.WriteLine "========== CARPETAS =========="
        ts.WriteBlankLines 1
        ts.Close
        
        For Each subCarpeta In fso.GetFolder(strCarpeta).SubFolders
            DoEvents
            If subCarpeta.Name <> "Listado_PC" Then
                Saltos = ""
                nivel = nivel + 1
            
                For n = 0 To nivel + 1
                    Saltos = Saltos & vbTab
                Next
                
                On Error Resume Next
                With subCarpeta
                    
                    Set ts = fl.OpenAsTextStream(ForAppending, TristateUseDefault)
                    
                    ts.WriteLine Saltos & "Nombre           : " & .Name
                    ts.Close
                    
                End With
                recorreCarpetas subCarpeta.Path
                nivel = nivel - 1
            End If
        Next
        If fso.GetFolder(strCarpeta).SubFolders.Count > 0 Then
            ts.WriteBlankLines 1
            ts.WriteLine "=============================="
        End If

     End If
    
End Sub


Function SacaAtributos(Atrib As Integer) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : SacaAtributos
' Autor original    : Alba Salvá
' Creado            : diferentes fechas
' Propósito         : extrayendo los atributos las diferentes carpetas. Se usa con la función "recorreCarpetas"
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Atrib         Obligatorio     Representa la atributo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Retorno           : devuelve una cadena con el atributo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub recorreCarpetas_test( unidadPadre)
'
'        Call recorrerPC
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim t As String, x As String
    
    If Atrib And 1 Then 'ReadOnly = 1
        t = "Sólo lectura"
    End If
    If Atrib And 2 Then 'Hidden = 2
        x = "Oculto"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 4 Then 'System = 4
        x = "Sistema"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 8 Then 'Volume = 8
        x = "Volumen"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 16 Then 'Directory = 16 ' (&H10)
        x = "Directorio"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 32 Then 'Archive = 32 ' (&H20)
        x = "Archivo"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 1024 Then 'Alias = 1024 ' (&H400)
        x = "Alias"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    If Atrib And 2048 Then 'Compressed = 2048 ' (&H800)
        x = "Comprimido"
        If t <> "" Then
            t = t & ", " & x
        Else
            t = x
        End If
    End If
    
    If t = "" Then
        t = "Normal" ' = 0
    End If
    
    sacaAtributos = t
    
End Function
