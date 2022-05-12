Option Compare Database
Option Explicit

'modVincenty
Private Const PI = 3.14159265358979
Private Const EPSILON As Double = 0.000000000001

Public Function distVincenty(ByVal lat1 As Double, ByVal lon1 As Double, ByVal Lat2 As Double, ByVal Lon2 As Double) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente             : https://access-global.net/vba-google-maps-api-calcular-la-distancia-entre-dos-coordenadas-vincenty
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título             : distVincenty
' Autor original     : Alba Salvá
' Fecha              : 21/02/2020
' Propósito          : Conocer la distancia geodésica entre dos puntos especificados por latitud/longitud usando la
'                      fórmula inversa de Vincenty para elipsoides
' Retorno            : devuelve la distancia en m (con una preción hasta milímetros)
' Argumento/s        : La sintaxis del procedimiento o función consta del siguiente argumento:
'                      Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                      LatitudInicio    Obligatorio    Latitud del punto 1
'                      LongitudInicio   Obligatorio    Longitud del punto 1
'                      LatitudFin       Obligatorio    Latitud del punto 2
'                      LongitudFin      Obligatorio    Longitud del punto s
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Mas información    : el código ha sido adaptado a VBA del javascript publicado en:
'                      http://www.movable-type.co.uk/scripts/latlong-vincenty.html
'                      fórmula inversa de Vincenty - T Vincenty, "Direct and Inverse Solutions of Geodesics on the
'                      Ellipsoid with application of nested equations", Survey Review, vol XXII no 176, 1975
'Referencia Adicional: http://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test                : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                      portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub distVincenty_test()
'
'Debug.print distVincenty(latitudorigen, longitudorigen, latituddestino, longituddestino)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

  Dim low_a As Double
  Dim low_b As Double
  Dim f As Double
  Dim L As Double
  Dim U1 As Double
  Dim U2 As Double
  Dim sinU1 As Double
  Dim sinU2 As Double
  Dim cosU1 As Double
  Dim cosU2 As Double
  Dim lambda As Double
  Dim lambdaP As Double
  Dim iterLimit As Integer
  Dim sinLambda As Double
  Dim cosLambda As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim sigma As Double
  Dim sinAlpha As Double
  Dim cosSqAlpha As Double
  Dim cos2SigmaM As Double
  Dim c As Double
  Dim uSq As Double
  Dim upper_A As Double
  Dim upper_B As Double
  Dim deltaSigma As Double
  Dim s As Double ' resultado final redondeado a 3 decimales (mm).
  
  Dim P1 As Double
  Dim P2 As Double
  Dim P3 As Double

'Ver http://es.wikipedia.org/wiki/World_Geodetic_System (en inglés)
'para information sobre los parámetros de varios Elipsoides de otros estándares.
'
'low_a y low_b en metros
' === GRS-80 ===
' low_a = 6378137
' low_b = 6356752.314245
' f = 1 / 298.257223563
'
' === Airy 1830 ===  Mayor precisión para Inglaterra y el norte de Europa
' low_a = 6377563.396
' low_b = 6356256.910
' f = 1 / 299.3249646
'
' === Internacional 1924 ===
' low_a = 6378388
' low_b = 6356911.946
' f = 1 / 297
'
' === Modelo Clarke 1880 ===
' low_a = 6378249.145
' low_b = 6356514.86955
' f = 1 / 293.465
'
' === GRS-67 ===
' low_a = 6378160
' low_b = 6356774.719
' f = 1 / 298.247167

'=== ParÃ¡metros Elipsoide WGS-84 === El más usado en todo el mundo, incluidos los sistemas GPS
  low_a = 6378137       ' +/- 2m
  low_b = 6356752.3142
  f = 1 / 298.257223563
'====================================
  L = toRad(Lon2 - lon1)
  U1 = Atn((1 - f) * Tan(toRad(lat1)))
  U2 = Atn((1 - f) * Tan(toRad(Lat2)))
  sinU1 = Sin(U1)
  cosU1 = Cos(U1)
  sinU2 = Sin(U2)
  cosU2 = Cos(U2)

  lambda = L
  lambdaP = 2 * PI
  iterLimit = 100 ' se puede disminuir hasta 20 si se desea.

  While (Abs(lambda - lambdaP) > EPSILON) And (iterLimit > 0)
    iterLimit = iterLimit - 1

    sinLambda = Sin(lambda)
    cosLambda = Cos(lambda)
    sinSigma = Sqr(((cosU2 * sinLambda) ^ 2) + ((cosU1 * sinU2 - sinU1 * cosU2 * cosLambda) ^ 2))
    If sinSigma = 0 Then
      distVincenty = 0  'puntos coincidentes
      Exit Function
    End If
    cosSigma = sinU1 * sinU2 + cosU1 * cosU2 * cosLambda
    sigma = Atan2(cosSigma, sinSigma)
    sinAlpha = cosU1 * cosU2 * sinLambda / sinSigma
    cosSqAlpha = 1 - sinAlpha * sinAlpha

    If cosSqAlpha = 0 Then 'verificamos di es divisiÃ³n por cero
      cos2SigmaM = 0 '2 puntos en el ecuador
    Else
      cos2SigmaM = cosSigma - 2 * sinU1 * sinU2 / cosSqAlpha
    End If

    c = f / 16 * cosSqAlpha * (4 + f * (4 - 3 * cosSqAlpha))
    lambdaP = lambda

'Los cálculos originales son muy complejos para VBA
'por ello, se han dividido en varias partes para evitar problemas.
'la implementación original para el cálculo de Lambda
'  lambda = L + (1 - C) * f * sinAlpha * _
            (sigma + C * sinSigma * (cos2SigmaM + C * cosSigma * (-1 + 2 * (cos2SigmaM ^ 2))))
      
    'calculamos porciones
      
    P1 = -1 + 2 * (cos2SigmaM ^ 2)
    P2 = (sigma + c * sinSigma * (cos2SigmaM + c * cosSigma * P1))
    
    'completo el cálculo
    lambda = L + (1 - c) * f * sinAlpha * P2

  Wend

  If iterLimit < 1 Then
    MsgBox "Se ha alcanzado el lÃ­mite de iteraciones," & vbCrLf & _
           "algo no ha idocomo se esperaba.", vbExclamation, "CÃ¡lculo por mÃ©todo Vincenty"
    Exit Function
  End If

  uSq = cosSqAlpha * (low_a ^ 2 - low_b ^ 2) / (low_b ^ 2)

  'Los cálculos originales son muy complejos para VBA
  'por ello, se han dividido en varias partes para evitar problemas.
  '
  'la implementación original para el cálculo de upper_A
  'upper_A = 1 + uSq / 16384 * (4096 + uSq * (-768 + uSq * (320 - 175 * uSq)))
  
  'calculo una parte de la ecuación
  P1 = (4096 + uSq * (-768 + uSq * (320 - 175 * uSq)))
  'completo el cÃ¡lculo
  upper_A = 1 + uSq / 16384 * P1

  'por extraño que parezca, upper_B calcula sin ningún problema
  upper_B = uSq / 1024 * (256 + uSq * (-128 + uSq * (74 - 47 * uSq)))

  'Los cálculos originales son muy complejos para VBA
  'por ello, se han dividido en varias partes para evitar problemas.
  '
  'la implementación original para el cálculo de deltaSigma
  'deltaSigma = upper_B * sinSigma * (cos2SigmaM + upper_B / 4 * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) _
                - upper_B / 6 * cos2SigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2)))
  
  'el cálculo de la fórmula de deltaSigma se divide en 3 partes
  'para prevenir el error de overflow que puede ocurrir
  
  P1 = (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2)
  P2 = upper_B * sinSigma
  P3 = (cos2SigmaM + upper_B / 4 * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) - upper_B / 6 * cos2SigmaM * P1))
  
  'completo el cálculo de deltaSigma
  deltaSigma = P2 * P3

  'calculo la distancia
  s = low_b * upper_A * (sigma - deltaSigma)
  
  'redondeo la distancia a milímetros
  distVincenty = Round(s, 3)

End Function

Function Convert_Degree(Decimal_Deg) As Variant
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : http://support.microsoft.com/kb/213449
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Convert_Degree
' Propósito         : converts a decimal degree representation to deg min sec
'                     as 10.46 returns 10° 27' 36"
' Retorno           : el valor en grados del valor que le pasamos
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     decimal_Deg     Obligatorio    Valor decimal obtenido
'-----------------------------------------------------------------------------------------------------------------------------------------------
  Dim degrees As Variant
  Dim minutes As Variant
  Dim seconds As Variant
  
  With Application
     'Set degree to Integer of Argument Passed
     degrees = Int(Decimal_Deg)
     'Set minutes to 60 times the number to the right
     'of the decimal for the variable Decimal_Deg
     minutes = (Decimal_Deg - degrees) * 60
     'Set seconds to 60 times the number to the right of the
     'decimal for the variable Minute
     seconds = Format(((minutes - Int(minutes)) * 60), "0")
     'Returns the Result of degree conversion
    '(for example, 10.46 = 10º 27' 36")
     Convert_Degree = " " & degrees & "º " & Int(minutes) & "' " _
         & seconds + Chr(34)
  
  End With

End Function

Function Convert_Decimal(Degree_Deg As String) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : http://support.microsoft.com/kb/213449
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Convert_Decimal
' Propósito         : Converts text angular entry to decimal equivalent, as:
'                     10° 27' 36" returns 10.46
'                     alternative to "°" is permitted: Use "~" instead, as:
'                     10~ 27' 36" also returns 10.46
' Retorno           : el valor en grados del valor que le pasamos
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     decimal_Deg     Obligatorio    Valor decimal obtenido
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Importante        : Declare the variables to be double precision floating-point.
'-----------------------------------------------------------------------------------------------------------------------------------------------
   Dim degrees As Double
   Dim minutes As Double
   Dim seconds As Double
   '
   
   '-----------------------------------------------------------------
   'modificación por JLatham
   'permite usar el símolo "~" symbol en vez de "°" para indicar grados
   'dado que "~" está disponible en teclados no espaañoles y "°" se tiene
   'que introducir por [Alt] [0] [1] [7] [6].
   Degree_Deg = Replace(Degree_Deg, "~", "°")
   '-----------------------------------------------------------------

   ' Set degree to value before "º" of Argument Passed.
   degrees = Val(Left(Degree_Deg, InStr(1, Degree_Deg, "º") - 1))
   
   ' Set minutes to the value between the "º" and the "'"
   ' of the text string for the variable Degree_Deg divided by
   ' 60. The Val function converts the text string to a number.
   minutes = Val(Mid(Degree_Deg, InStr(1, Degree_Deg, "º") + 2, _
             InStr(1, Degree_Deg, "'") - InStr(1, Degree_Deg, "º") - 2)) / 60
   
   ' Set seconds to the number to the right of "'" that is
   ' converted to a value and then divided by 3600.
   seconds = Val(Mid(Degree_Deg, InStr(1, Degree_Deg, "'") + _
           2, Len(Degree_Deg) - InStr(1, Degree_Deg, "'") - 2)) / 3600
   
   Convert_Decimal = degrees + minutes + seconds

End Function

Private Function toRad(ByVal degrees As Double) As Double
    
    toRad = degrees * (PI / 180)

End Function

Private Function Atan2(ByVal X As Double, ByVal Y As Double) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : http://en.wikibooks.org/wiki/Programming:Visual_Basic_Classic/Simple_Arithmetic#Trigonometrical_Functions
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Atan2
' Propósito         : Converts text angular entry to decimal equivalent, as:
'                     10° 27' 36" returns 10.46
'                     alternative to "°" is permitted: Use "~" instead, as:
'                     10~ 27' 36" also returns 10.46
' Retorno           : la arcotangente de las coordenadas que le pasamos a la función
' Argumento/s       : La sintaxis del procedimiento o función consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     decimal_Deg     Obligatorio    Valor decimal obtenido
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Importante        : Si reutilizas este código, ten en cuenta que X e Y se han invertido respecto al uso tópico
'-----------------------------------------------------------------------------------------------------------------------------------------------

    If Y > 0 Then
        If X >= Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= -Y Then
            Atan2 = Atn(Y / X) + PI
        Else
        Atan2 = PI / 2 - Atn(X / Y)
    End If
        Else
            If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - PI
        Else
            Atan2 = -Atn(X / Y) - PI / 2
        End If
    End If

End Function

