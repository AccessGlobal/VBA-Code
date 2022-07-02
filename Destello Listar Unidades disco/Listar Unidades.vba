Sub ListaUnidadesDisco()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-listar-unidades-de-disco
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListaUnidadesDisco
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Listar todas las unidades de disco y sus atributos
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/es-es/office/vba/language/reference/user-interface-help/drive-object
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim Discos As Object
Dim Disco As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        Set Discos = fso.Drives
        
        Debug.Print "Tu PC tiene " & Discos.Count & " unidades de disco"
        
        For Each Disco In Discos
            DoEvents
            With Disco
                Debug.Print vbTab; .DriveLetter & " - ";
                Select Case .DriveType
                    Case 1
                        Debug.Print "Removible";
                    Case 2
                        Debug.Print "Fijo";
                    Case 3
                        Debug.Print " en Red";
                    Case 4
                        Debug.Print "CDRom";
                    Case 5
                        Debug.Print "Disco RAM";
                    Case Else
                        Debug.Print "Desconocido";
                End Select
                Debug.Print "; ";
                If .DriveType = 3 Then
                    Debug.Print "Recurso de red: "; .ShareName
                Else
                    If .IsReady Then
                        Debug.Print "Nombre: "; .VolumeName
                    Else
                        Debug.Print "Unidad no disponible"
                    End If
                End If
                
                Debug.Print vbTab; vbTab; "Está activo: "; .IsReady
                
                If .IsReady Then
                    Debug.Print vbTab; vbTab; "Nº de serie: "; .SerialNumber
                    Debug.Print vbTab; vbTab; "Sistema de srchivos: "; .FileSystem
                    Debug.Print vbTab; vbTab; "Capacidad total: "; Format(.TotalSize, "#,##0"); " bytes"
                    Debug.Print vbTab; vbTab; "Espacio libre  : "; Format(.FreeSpace, "#,##0"); " bytes"
                    Debug.Print vbTab; vbTab; "Carpeta raiz: "; .RootFolder
                End If
                
                Debug.Print vbTab; vbTab; "Ruta: "; .Path
            
            End With
        Next
    Set fso = Nothing
    
End Sub