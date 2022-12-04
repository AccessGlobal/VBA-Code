'Crear un módulo estándar 
Option Compare Database
Option Explicit

Public Const API_KEY = "API KEY"

Public Function GetpriceCrypto(symbol As String) As Single
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-el-fascinante-mundo-de-las-api
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetpriceCrypto
' Autor original    : Luis Viadel
' Fecha             : mayo 2021
' Propósito         : Obtener la cotización de diversas Criptomonedas en tiempo real mediante la llamada a la API de CoinMarketcap
' Retorno           : Valor single con la cotización de la criptomoneda que le pasamos en la variable symbol
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     symbol        Obligatorio    string que contiene el código de la criptomoneda de la que se dese conocer su cotización
'                                                  Bitcoin --> BTC, Ethereum --> ETH, ...
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : c:\window\SysWoW64\smsxml6.dll
' Más información   : https://coinmarketcap.com/api/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub GetpriceCrypto_test()
' Dim BTCValue as single
'
'    BTCvalue=GetpriceCrypto("BTC")
'
'    Debug.Print BTCValue
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objXMLHTTP As MSXML2.XMLHTTP60
Dim sURL As String, str1 As String
Dim cadenaTemp As String
Dim IntWhere As Long

    Select Case symbol
        
        Case "ETH"
            Set objXMLHTTP = New MSXML2.XMLHTTP60
                       
                sURL = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=" & symbol
                    
                With objXMLHTTP
                    .Open "GET", sURL, False
                    .setRequestHeader "X-CMC_PRO_API_KEY", API_KEY
                    .setRequestHeader "Accepts", "application/json"
                    .Send ("")
                End With
                              
                str1 = objXMLHTTP.responseText
                
                Debug.Print str1
                
            Set objXMLHTTP = Nothing
            
            IntWhere = InStr(1, str1, """USD""" & ":{" & """price""" & ":")
            
            cadenaTemp = right(str1, Len(str1) - IntWhere - 14)
            
            cadenaTemp = left(cadenaTemp, 14)
            
            cadenaTemp = Replace(cadenaTemp, ".", ",")
                   
            GetpriceCrypto = CSng(cadenaTemp)
            
    End Select
    
End Function
