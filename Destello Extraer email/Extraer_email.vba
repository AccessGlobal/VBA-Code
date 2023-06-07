Public Function ExtractEmailAddress(strData As String, _
    Optional strDelim As String = ",") As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-extraer-email-de-una-cadena-de-texto/
'                     Destello formativo 336
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExtractEmailAddress
' Autor original    : thedbguy | http://www.accessmvp.com/thedbguy
' Creado            : junio 2015
' Propósito         : extraer una dirección de mail de una cadena de texto
' Argumento         : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strData        Obligatorio      cadena de texto que contiene el email
'                     strDelim        Opcional        delimitador de cada bloque de texto en la cadena
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://www.regular-expressions.info/email.html
'                     https://www.robvanderwoude.com/vbstech_regexp.php
'                     https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Private Sub btnQRTest_Click()
'Dim resultado As String
'
'    resultado = ExtraeEmail("Pedro Martínez, C/calle de pedro 23, 2, 46015 Valencia, pedro@webdepedro.com, 666 666 666")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim regEx As Object
Dim regExMatch As Object
Dim var As Variant
Dim strEmail As String

    Set regEx = CreateObject("VBScript.RegExp")
    
        With regEx
'Three properties are available for the RegExp object:
'Global: if TRUE, find all matches, if FALSE find only the first match
            .Global = True
'IgnoreCase: if TRUE perform a case-insensitive search, if FALSE perform a case-sensitive search
            .IgnoreCase = True
'Pattern: the RegExp pattern to search for
            .Pattern = "\b[0-9A-Z._%+-]+@[0-9A-Z.-]+.[A-Z]{2,3}\b"
'Execute method: returns an object with the following properties:
' Count: the number of matches found in teststring (maximum is 1 if .Global = False)
' Item: the matches themselves as objects, each with the following properties:
'       FirstIndex: the location of the matching substring in teststring
'       Length    : the length of the matching substring
'       SubMatches: if parentheses were used in the pattern: the matching pieces of the pattern between sets of parentheses as objects, each with the following properties:
'                   Count: the number of submatches found in the match
'                   Item : the string value of the submatch
'                   value: the string value of the match
'Another methods: test, replace
            Set regExMatch = .Execute(strData)
                For Each var In regExMatch
                    strEmail = strEmail & strDelim & var
                Next
            Set regExMatch = Nothing
       End With
        
        ExtractEmailAddress = Mid(strEmail, Len(strDelim) + 1)
    
    Set regEx = Nothing

End Function
