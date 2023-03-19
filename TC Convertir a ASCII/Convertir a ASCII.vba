
Public Function mcstrConvertToAscii(ByVal strString As String, ByVal blnMantenerFormatoMayúscula As Boolean, Optional ByVal blnMantenerEspacios As Boolean, Optional ByVal strCarácterAObviar As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/funcion-que-convierte-cualquier-caracter-a-texto-puramente-ascii-access-vba
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcstrConvertToAscii
' Autor             : Rafael .:McPegasus:. Copyright ©1999-2007 for Puzzle
' Actualizado       : 20/07/2021
' Propósito         : Pasar una cadena de texto que contiene acentos, tildes, acentos circunflejos (palabras francesas), eñes, diéresis (común en el alemán), cedillas y otros dicríticos, a texto puramente ASCII.
' Retorno           : Una cadena de texto con sólo caracteres ASCII.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Argumentos        : La sintaxis del procedimiento o función consta de los siguientes argumentos:
'                     Parte                 Modo           Descripción
'                     --------------------------------------------------------------------------------------------------------------------------
'                     strString             Obligatorio    El valor String especifica una cadena de texto que contiene acentos, acentos circunflejos (palabras francesas), eñes, diéresis (común en el alemán), cedillas y otros dicríticos.
'                     blnMantenerFormatoMayúscula  Opcional El valor Boolean especifica si se desea mantener en mayúscula los caracteres que así estén de origen.
'                     [ blnMantenerEspacios ]  Opcional    El valor Boolean especifica si se desea mantener los espaciones o eliminarlos.
'                     [ strCarácterAObviar ]  Opcional     El valor String especifica si se desea no sustituir algún caracter.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Sobre Referenciar : El referenciar una librería externa nos permite seleccionar los objetos de otra aplicación que se desea que estén disponibles en nuestro código. También acceder a sus métodos utilizar las constantes.
'                     En caso de ser opcional podemos seguir utilizándolo aunque las constantes hay que sustituirlas por su valor, normalmente numérico.
'                     Más información: https://support.microsoft.com/es-es/office/add-object-libraries-to-your-visual-basic-project-ed28a713-5401-41b0-90ed-b368f9ae2513
' Referencia        : Opcional. Microsoft Scripting Runtime (c:\Windows\SysWOW64\scrrun.dll)
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar todo el procedimiento desde el Sub hasta el End Sub
'                     al portapapeles y pega en el editor de VBA de tu aplicación MS Access. Descomentar todas las líneas que nos interese (se aconseja seleccionar
'                     todas las líneas del ejemplo y utilizar el botón 'Bloque sin comentarios' de la barra de herramientas 'Edición').
'                     Pulsar F5 para ver su funcionamiento.
'
'    Sub mcstrConvertToAscii_test()
'
'        Dim strCadena                               As String
'
'
'        strCadena = "€ÍCœ€amión€ÓáÁÉé-"
'
'        Debug.Print
'        Debug.Print "Original: " & strCadena
'        Debug.Print "Mayúscu.: " & mcstrConvertToAscii(strCadena, True)               'Mantener las mayúsculas.
'        Debug.Print "Minúscu.: " & mcstrConvertToAscii(strCadena, False)              'Convertir a minúsculas.
'
'    End Sub
'</Test>
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Importante        : A comienzo en el módulo, comprobar que está la declaración "Option Compare Binary" para que el código distinga entre minúsculas y mayúsculas.
'-----------------------------------------------------------------------------------------------------------------------------------------------

