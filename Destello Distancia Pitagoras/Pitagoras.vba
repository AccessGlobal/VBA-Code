Option Compare Database
Option Explicit

Function DistPitagoras(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-google-maps-api-calcular-la-distancia-entre-dos-coordenadas-pitagoras
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : DistPitagoras
' Autor original    : Alba Salvá
' Fecha             : 1980 - 1982
' Propósito         : Conocer la distancia entre dos puntos geográficos suponiendo que ambos puntos están en un plano, lo que nos permite
'                     utilizar el conocido "Teorema de Pitágoras".
' Retorno           : devuelve la distancia en km
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     LatitudInicio    Obligatorio    Latitud del punto 1
'                     LongitudInicio   Obligatorio    Longitud del punto 1
'                     LatitudFin       Obligatorio    Latitud del punto 2
'                     LongitudFin      Obligatorio    Longitud del punto 2
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test               : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub DistPitagoras_test()
'
'Debug.print DistPitagoras(latitudorigen, longitudorigen, latituddestino, longituddestino)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Dist As Double
Dim P1 As Double, P2 As Double

P1 = (Lat1 - Lat2) ^ 2
P2 = (Lon1 - Lon2) ^ 2

DistPitagoras = (Sqr(P1 + P2)) * 100

End Function


'===============================================================================================================================================
'===============================================================================================================================================

' Versión mejorada con resultados con una mayor precisión.
' Esta modificación fue realizada en torno a los años 1985 -1987.
' Los parámetros de entrada son los mismos, la diferencia con la versión anterior estriba en que se incluyen cálculos modificadores para las distancias, 
' en vez de usar un valor fijo.

Function DistPitagoras(Lat1 As Double, Lon1 As Double, Lat2 As Double, Lon2 As Double) As Double

Dim Dist As Double
Dim P1 As Double, P2 As Double
Dim M As Double, Z As Double

Const PI = 3.14159265358979
Const D As Double = 111.12 'Distancia en Km de 1 grado de latitud en el ecuador (aproximado)

P1 = ((Lat1 - Lat2) * D) ^ 2 'Calculamos el cuadrado de la distancia entre latitudes

M = ((Lat1 + Lat2) / 2) * (PI / 180)'Convertimos a Radianes la media de las latitudes
Z = Cos(M) * D ' Calculamos el corrector para la distancia de las longitudes

P2 = ((Lon1 - Lon2) * Z) ^ 2 'Calculamos el cuadrado de la distancia de las longitudes

DistPitagoras = (Sqr(P1 + P2)) 'Devolvemos el resultado con la fórmula de Pitágoras

End Function

