'Declarar a nivel de módulo
Option Compare Database
Option Explicit

'Variables para capturar las dimensiones del formulario de base
Public Ancho As Single
Public alto As Single
'Variables para conocer las dimensiones del formulario antes de redimensionar
Public anchoformposicion As Single
Public altoformposicion As Single
'Variables para guardar los parámetros del interior del formulario base
Public Dim1X As Integer
Public Dim1Y As Integer
'Variables para la guardar la posición X e Y del formulario de base
Public PosX As Integer
Public PosY As Integer

Public Function posForm(frm As Form)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-redimensionar-y-posicionar-formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : posForm
' Autor original    : Luis Viadel
' Actualizado       : febrero 2015
' Propósito         : Conocer la posición de cualquier formulario en la pantalla.
' Retorno           : sin retorno
' Argumento/s       : La sintaxis de la función consta de los siguientes argumentos:
'                     Parte                 Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm              Obligatorio        Nombre del formulario de que queremos saber su posición
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub PosForm_test()
'
'    Call PosForm(Me)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------Dim mu11 As Single, mul2 As Single

PosX = frm.Properties!WindowLeft
PosY = frm.Properties!WindowTop
    
Dim1X = frm.Width
Dim1Y = frm.Section(acDetail).Height

Form_FormPosicion.InsideWidth = Dim1X
Form_FormPosicion.InsideHeight = Dim1Y

'Calculamos los multiplicadores para poder adaptar los controles
mu11 = (Dim1X / anchoformposicion) + 0.02
mul2 = (Dim1Y / altoformposicion) + 0.02

'Adaptamos los controles según los multiplicadores
Form_FormPosicion.CuadroPuntos.Width = Form_FormPosicion.CuadroPuntos.Width * mu11
Form_FormPosicion.CuadroPuntos.Height = Form_FormPosicion.CuadroPuntos.Height * mul2

'Si hubiesen más controles, por ejemplo un cuadro de texto, podríamos aumentar el tamaño de letra multiplicando
'.fontsize por el multiplicador

End Function

