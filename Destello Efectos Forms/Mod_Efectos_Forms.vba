'Código a incorporar en un módulo estándar
Option Compare Database
Option Explicit

' Módulo Efectos de ventana
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-efectos-ventana/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : modEfectos
' Autor original    : Luis Viadel | @luisviadel | https://cowtechnologies.net
' Fecha             : junio 12
' Propósito         : configurar todas las variables necesarias y las funciones para poder crear efectos de cierre y apertura utilizando
'                     la API de Windows con la función AnimateWindow
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Recursos externos : No son necesarios
' Importante        : el efecto se puede manejar a traves del tiempo de ejecución del mismo que lo marca la variable "TiempoAnimacion".
'                     La función AnimateWindow le permite producir efectos especiales al mostrar u ocultar ventanas. Hay tres tipos de
'                     animación: roll, slide y alpha blended fade.
' Más información   : https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-animatewindow
' Test para generar efecto
'      Sub GeneraEfecto_test(hWnd As Long, AnimationTime As Long, flag As Long)
'            AnimateWindow hWnd, AnimationTime, flag
'      End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public Declare PtrSafe Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long

'Se pueden utilizar con animación rollo o de diapositiva
Public Const AW_HOR_POSITIVE = &H1   'La animación de la ventana se produce de izquierda a derecha
                

Public Const AW_HOR_NEGATIVE = &H2   'La animación de la ventana se produce de derecha a izquierda


Public Const AW_VER_POSITIVE = &H4   'La animación de la ventana se produce desde arriba hacia abajo


Public Const AW_VER_NEGATIVE = &H8   'La animación de la ventana se produce desde abajo hacia arriba


Public Const AW_CENTER = &H10        'Hace que la ventana parezca contraerse hacia adentro si se usa AW_HIDE _
                                      o se expande hacia afuera si no se usa AW_HIDE.
               
Public Const AW_HIDE = &H10000       'Oculta la ventana. Por defecto la ventana es visible.

Public Const AW_ACTIVATE = &H20000   'Activa la ventana.

Public Const AW_SLIDE = &H40000      'Utiliza animación de diapositivas. De forma predeterminada, se utiliza la animación de rollo.

Public Const AW_BLEND = &H80000      'Utiliza un efecto de desvanecimiento. Esta bandera solo se puede usar si hwnd es una ventana de nivel superior.

Public Const TiempoAnimacion = 700   'Tiempo de la animación en milisegundos


Sub GeneraEfecto(hWnd As Long, AnimationTime As Long, flag As Long)
  
  AnimateWindow hWnd, AnimationTime, flag

End Sub

