Option Compare Database
Option Explicit

Public Declare PtrSafe Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fwinIni As Long) As Long

Public Const SPI_GETWORKAREA = 48

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private DimX As Long, DimY As Long
Private Dim1X As Long, Dim1Y As Long
Private PosX As Long, PosY As Long
Dim ancho As Long, alto As Long

Public Function posForm(frm As Form)
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-modificar-o-eliminar-origen-de-datos-en-tiempo-de-ejecucion/
'                      Destello formativo 334
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : PosForm
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : 21/04/2015
' Propósito         : Conocer la posición de cualquier formulario en la pantalla.
'                     Siempre se ejecuta en primer lugar. Fomulario que hace la llamada.
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm            Obligatorio      primer formulario que realiza la llamada
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub PosForm_test()
'
'      Call PosForm(Me)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ret As Long
Dim T_rect As RECT

    ret = SystemParametersInfo(SPI_GETWORKAREA, 0, T_rect, 0)
    
'Dimensiones de la pantalla para evitar que el formulario emergente se salga de la misma
    ancho = (T_rect.Right - T_rect.Left) * 15
    alto = (T_rect.Bottom - T_rect.Top) * 15

    PosX = frm.Properties!WindowLeft
    PosY = frm.Properties!WindowTop
    
    Dim1X = frm.Width
    Dim1Y = frm.Section(acDetail).Height

End Function

Public Function DimForm(frm As Form)
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-modificar-o-eliminar-origen-de-datos-en-tiempo-de-ejecucion/
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : DimForm
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : 21/04/2015
' Propósito         : Conocer las dimensiones de cualquier formulario.
'                     Siempre se ejecuta en segundo lugar. Fomulario que recibe la llamada.
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm            Obligatorio      formulario que deseamos posicionar
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub DimForm_test()
'
'      Call DimForm(Nombre del formulario)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    DimX = frm.Width
    DimY = frm.Section(acDetail).Height
   
   Call UbicaForm(frm)

End Function

Public Function UbicaForm(frm As Form)
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-modificar-o-eliminar-origen-de-datos-en-tiempo-de-ejecucion/
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : UbicaForm
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : 21/04/2015
' Propósito         : Ubicar correctamente el formulario si se sale de la pantalla
' Argumentos        : la sintaxis de la función consta un argumento:
'                     Parte             Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm            Obligatorio      formulario que deseamos posicionar
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub UbicaForm_test()
'
'      Call UbicaForm(Nombre del formulario)
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

    On Error Resume Next
    
    If ancho < (PosX + Dim1X + DimX) Then
        frm.Move Left:=PosX - (DimX / 2), Top:=PosY + DimY
        frm.SetFocus
    Else
        frm.Move Left:=PosX + Dim1X - (DimX / 2), Top:=PosY + (DimY / 2)
        frm.SetFocus
    End If

End Function
