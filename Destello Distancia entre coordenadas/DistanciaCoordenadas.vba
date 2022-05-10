Option Compare Database
Option Explicit

Private Const PI As Double = 3.14159265358979
'el valor del radio ecuatorial de la tierra es de 6378 km
'El radio polar de la tierra es de 6357 km.
'El radio equivolumen es de 6371 km.
Private Const RadioTerrestre As Double = 6378

Function CalculoKM(LatitudInicio, LongitudInicio, LatitudFin, LongitudFin) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-google-maps-api-calcular-la-distancia-entre-dos-coordenadas
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CalculoKM
' Autor original    : Luis Viadel
' Fecha             : 21/02/2020
' Propósito         : Conocer la distancia entre dos puntos geográficos. Para hacerlo, vamos a utilizar la "fórmula de Haversine" que, sin entrar
'                     en detalles matemáticos, es la siguiente:
'                          R = radio de la Tierra
'
'                          DiferenciaLatitud = lat2- lat1
'
'                          DiferenciaLongitud = long2- long1
'
'                          a = sin²(DiferenciaLatitud/2) + cos(lat1) · cos(lat2) · sin²(DiferenciaLongitud/2)
'
'                          c = 2 · atan2(va, v(1-a))
'
'                          d = R · c
'
' Retorno           : devuelve la distancia en km
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     LatitudInicio    Obligatorio    Latitud del punto 1
'                     LongitudInicio   Obligatorio    Longitud del punto 1
'                     LatitudFin       Obligatorio    Latitud del punto 2
'                     LongitudFin      Obligatorio    Longitud del punto s
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Mas información   : http://es.wikipedia.org/wiki/F%C3%B3rmula_del_Haversine
' Importante        : La  Fórmula del Haversine es de las más utilizadas para el cálculo de distancias entre dos puntos (hay otras como la Ley
'                     Esférica del Coseno), pero asume que la tierra es una esfera perfecta y no lo es, por lo que los cálculos están sujetos a
'                     error. Si quieres seguir investigando para aumentar la fiabilidad, puedes empezar por aquí:
'                     http://www.movable-type.co.uk/scripts/gis-faq-5.1.html
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test               : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub calculoKM_test()
'
'Debug.print CalculoKM(latitudorigen, longitudorigen, latituddestino, longituddestino)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim constante As Double
Dim DiferenciaLatitud As Double
Dim DiferenciaLongitud As Double
Dim A As Double
Dim ConstRadianes As Double

'Las latitudes y las longitudes se indican en grados, minutos y segundos por lo que las debemos pasar a radianes.
ConstRadianes = PI / 180

'Calculamos los incrementos de latitud y longitud
DiferenciaLatitud = LatitudInicio - LatitudFin
DiferenciaLongitud = LongitudInicio - LongitudFin

A = Sin(DiferenciaLatitud * ConstRadianes / 2) ^ 2 + Cos(LatitudInicio * ConstRadianes) * Cos(LatitudFin * ConstRadianes) * Sin(DiferenciaLongitud * ConstRadianes / 2) ^ 2

'Access no dispone de una función Arcoseno nativa, por lo que tenemos que realizar el cálculo matemático a través de una función
A = 2 * ArcSin(Sqr(A))

CalculoKM = A * RadioTerrestre

End Function

Function ArcSin(X As Double) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-freefile-function
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ArcSin
' Autor original    : Francisco Megía
' Fecha             : 20/07/2011
' Propósito         : cálculo del arcoseno
' Retorno           : Devuelve la el valor del arcoseno
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                       X       Obligatorio     Valor del ángulo del que queremos calcular su arcoseno
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Más información    : https://www.necesitomas.com/funciones_VBA_derivadas
'-----------------------------------------------------------------------------------------------------------------------------------------------

    ArcSin = Atn(X / Sqr(-X * X + 1))

End Function

