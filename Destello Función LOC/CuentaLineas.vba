Public Function CuentaLineas(strArchivo As String) As Long
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funcion-loc/
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CuentaLineas
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : mayo 2023
' Propósito         : 1. Contar las líneas de un fichero de texto
'                     2. Recuperar la cantidad de texto que se desee
' Argumentos        : La sintaxis de la función consta de los siguientes argumentos
'                     Variable          Modo          Descripción
'--------------------------------------------------------------------------------------------------------------------------------------------
'                     strArchivo      Obligatorio    path completo del archivo que se quiere analizar
' Retorno           : valor long con el número de líneas que contiene el fichero
'---------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/loc-function
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pégalo en el editor de VBA. Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
' Sub CuentaLineas_test()
' Dim strArchivo As string
' Dim resultado as long
'
'   strArchivo = "Path_completo_de_mi_archivo"
'
'      Resultado=CuentaLineas(strArchivo)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim numLinea As Long
Dim strLinea As String
    
    Open strArchivo For Binary As #1
    
    Do While numLinea < LOF(1)
        strLinea = strLinea & Input(1, #1)
        numLinea = Loc(1)
'Mostramos los cien primeros caracteres
'        If numLinea = 100 Then
'            Debug.Print strLinea
'            GoTo LinCLose
'        End If
    Loop
'Contamos el número de caracteres
    CuentaLineas = numLinea
'    CuentaLineas= LOF(1)
LinCLose:
    Close #1

End Function
