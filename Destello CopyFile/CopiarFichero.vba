Public Function CreaCopiaTemp(strRutafichero As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-copyfile
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CreaCopiaTemp
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         :
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/copyfile-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub CreaCopiaTemp_test()
' Dim ruta As String
'
'    ruta = "C:\MiPrograma\MiApp.accdb"
'
'    Debug.Print CreaCopiaTemp(ruta)
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim DbTemp As String
Dim nombre As String
Dim nmat As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
'Creamos la carpeta temporal para grabar la copia
        MkDir "C:\temp\"
'Extraemos el nombre
        nmat = Split(strRutafichero, "\")
        nombre = nmat(UBound(nmat))
'Construímos un nuevo nombre añadiendo temp al fichero
        nombre = "temp_" & nombre
'Construímos el path con el método BuildPath
        DbTemp = fso.BuildPath("C:\temp\", nombre)
'Copiamos el fichero en el fichero temporal
        fso.CopyFile strRutafichero, DbTemp, True
        
        CreaCopiaTemp = True
    
        GoTo lbFinally

lbError:
        If Err = 0 Then
    
        Else
            MsgBox Err & vbCrLf & Err.Description
        End If
    
lbFinally:
        On Error GoTo 0
    Set fso = Nothing
    
End Function
