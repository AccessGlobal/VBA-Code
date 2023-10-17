Sub JoinArrayTest()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funcion-join
'                     Destello formativo 363
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : JoinArrayTest
' Autor original    : Luis Viadel
' Fuente original   : Luis Viadel | luisviadel@access-global.net
' Creado            : 2023
' Propósito         : probar el funcionamiento de la función "join"
' Mas información   : https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/join-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim miMatriz() As String
Dim Miruta As String

    ReDim miMatriz(3)
    
    miMatriz(0) = "C:"
    miMatriz(1) = "Mis documentos"
    miMatriz(2) = "Documentos de gestión"
    
    Miruta = Join(miMatriz, "/")
    
    Debug.Print Miruta
    
End Sub