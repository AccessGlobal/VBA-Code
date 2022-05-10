Public Function cpMap(ByVal Direc As String, ByVal POB As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-conocer-codigo-postal-y-coordenadas-con-la-api-de-google-maps
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : cpMap
' Autor original    : Luis Viadel
' Fecha             : 13/01/2020
' Propósito         : Búsqueda de códigos postales y coordenadas en Google Maps
' Retorno           : Devuelve las coordenadas de la posición y el C.P.
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     direc      Obligatorio     Dirección a buscar
'                      pob       Obligatorio     Población
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : Microsoft XML, v6.0
' Mas información   : https://mapsplatform.google.com/
' Importante        : deberemos crear una cuenta de desarrollador para obtener una APIKey
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub cmpa_test()
'Dim txtzip as string
'Dim direc as string, pob as string
'
'direc=""
'pob=""
'txtzip = cpMap(Direc, POB)
'Debug.print txtZip

'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objXMLHTTP As MSXML2.XMLHTTP60
Dim sURL As String, datos As String, Str1 As String, subdata As String, CP As String
Dim J As Integer, I As Integer, intWhere As Integer
Dim ArrayRes

On Error Resume Next

Set objXMLHTTP = New MSXML2.XMLHTTP60
        
    sURL = "https://maps.googleapis.com/maps/api/place/textsearch/json?query="
    sURL = sURL & UTF8(Direc) & "%20" & UTF8(POB)
    sURL = sURL & "&key=" & APIKey
        
    With objXMLHTTP
        .Open "GET", sURL, False
        .setRequestHeader "Content-Type", "application/json"
        .Send ("")
    End With
                  
    Str1 = objXMLHTTP.responseText
Debug.Print Str1
Set objXMLHTTP = Nothing

    ArrayRes = Split(Str1, "formatted_address")

    J = UBound(ArrayRes) - LBound(ArrayRes) + 1

    For I = 1 To 1
        subdata = ArrayRes(I)
        datos = Right(subdata, Len(subdata) - 5)
        intWhere = InStr(datos, "geometry")
        datos = Left(datos, intWhere - 1)
        
        intWhere = InStr(datos, ",") 'Calle
        datos = Right(datos, Len(datos) - intWhere - 1)
        
        intWhere = InStr(datos, ",")
        datos = Right(datos, Len(datos) - intWhere - 1)
        CP = Left(datos, 5)
    Next I

    ArrayRes = Split(Str1, "lat")

    J = UBound(ArrayRes) - LBound(ArrayRes) + 1

    For I = 1 To 1
        subdata = ArrayRes(I)
        datos = Right(subdata, Len(subdata) - 4)
        intWhere = InStr(datos, ",")
        latitud = Left(datos, intWhere - 1)
        latitud = Trim(latitud)
    Next I
    
    ArrayRes = Split(Str1, "lng")
    
    J = UBound(ArrayRes) - LBound(ArrayRes) + 1

    For I = 1 To 1
        subdata = ArrayRes(I)
        datos = Right(subdata, Len(subdata) - 4)
        intWhere = InStr(datos, "}")
        longitud = Left(datos, intWhere - 1)
        longitud = Trim(longitud)
   Next I

cpMap = CP & "," & latitud & "," & longitud

End Function

Function UTF8(strTexto As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-conocer-codigo-postal-y-coordenadas-con-la-api-de-google-maps
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : UTF8
' Autor original    : desconocido
' Fecha             : desconocida
' Propósito         : modificar cadenas de texto para que sean legibles por un navegador web
' Retorno           : Devuelve la cadena transformada
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strTexto   Obligatorio     Cadena que queremos trnsformar
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub cmpa_test()
'Dim strtxto as string
'
'strtxto=""
'strtxto = UTF8(strtxto)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
strTexto = Replace(strTexto, " ", "%20")
strTexto = Replace(strTexto, "Ñ", "%C3%91")
strTexto = Replace(strTexto, "ñ", "%C3%B1")
strTexto = Replace(strTexto, "á", "%C3%A1")
strTexto = Replace(strTexto, "à", "%C3%A0")
strTexto = Replace(strTexto, "â", "%C3%A2")
strTexto = Replace(strTexto, "ã", "%C3%A3")
strTexto = Replace(strTexto, "ä", "%C3%A4")
strTexto = Replace(strTexto, "å", "%C3%A5")
strTexto = Replace(strTexto, "è", "%C3%A8")
strTexto = Replace(strTexto, "é", "%C3%A9")
strTexto = Replace(strTexto, "ê", "%C3%AA")
strTexto = Replace(strTexto, "ë", "%C3%AB")

strTexto = Replace(strTexto, "ì", "%C3%AC")
strTexto = Replace(strTexto, "í", "%C3%AD")
strTexto = Replace(strTexto, "î", "%C3%AE")
strTexto = Replace(strTexto, "ï", "%C3%AF")

strTexto = Replace(strTexto, "ð", "%C3%B0")
strTexto = Replace(strTexto, "ò", "%C3%B2")
strTexto = Replace(strTexto, "ó", "%C3%B3")
strTexto = Replace(strTexto, "ô", "%C3%B4")
strTexto = Replace(strTexto, "õ", "%C3%B5")
strTexto = Replace(strTexto, "ö", "%C3%B6")

strTexto = Replace(strTexto, "ù", "%C3%B9")
strTexto = Replace(strTexto, "ú", "%C3%BA")
strTexto = Replace(strTexto, "û", "%C3%BB")
strTexto = Replace(strTexto, "ü", "%C3%BC")

strTexto = Replace(strTexto, ",", "%2C")
strTexto = Replace(strTexto, "ý", "%C3%BD")
strTexto = Replace(strTexto, "þ", "%C3%BE")
strTexto = Replace(strTexto, "ÿ", "%C3%BF")

strTexto = Replace(strTexto, "÷", "%C3%B7")

UTF8 = strTexto

End Function