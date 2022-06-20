Sub MuestraProfundidadRuta(strCarpeta)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-getfolder-IsRootFolder-ParentFolder
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : MuestraProfundidadRuta
' Autor             : Alba Salvá
' Fecha             : desconocida
' Propósito         : Devuelve el nivel de profundidad de una carpeta.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/getfolder-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que interese y pulsar F5 para ver su funcionamiento.
'
' Sub MuestraProfundidadRuta_test()
' Dim ruta As String
'
'    ruta = "C:\MiPrograma\subcarpeta1\subcarpeta2"
'
'    Debug.Print MuestraProfundidadRuta(ruta)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim fso As Object
Dim f As Object
Dim n As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
                   
        Set f = fso.GetFolder(strCarpeta)
        
            If f.IsRootFolder Then
                MsgBox "La carpeta especificada es la carpeta raíz"
            Else
                Do Until f.IsRootFolder
                    Set f = f.ParentFolder
                    n = n + 1
                Loop
                MsgBox "La carpeta especificada está a " & n & " niveles por debajo."
            End If
        
        Set f = Nothing
    
    Set fso = Nothing
    
End Sub
