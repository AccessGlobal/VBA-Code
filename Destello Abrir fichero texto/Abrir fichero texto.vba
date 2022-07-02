Public Function AbrirFicheroTexto(nomfichero As String, Optional modo As Integer, Optional crear As Boolean, Optional estado As integer) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-metodo-opentextfile
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AbrirFicheroTexto
' Autor             : Luis Viadel
' Fecha             : noviembre 2019
' Propósito         : abrir un fichero de texto para escribir valores en él
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub AbrirFiceroTexto_test ()
'
'   Debug.Print AbrirFicheroTexto ("C:\MiFichero.txt",8,True, -2)
'
' End sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim fso As Object, fs As Object
    
    On Error GoTo LinErr
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        
        Set fs = fso.OpenTextFile(nomfichero, modo, crear, estado)
            
            fs.Write "Nueva línea de escritura"
            
            fs.Close
            
            AbrirFicheroTexto = True
        
        Set fs = Nothing
    
    Set fso = Nothing
        
    Exit Function
    
LinErr:
    AbrirFicheroTexto = False

End Function