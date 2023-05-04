Public Function GetTextLength(pCtrl As Control, ByVal str As String, Optional ByVal Height As Boolean = False)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/wizhook-series-adaptar-un-cuadro-de-texto-a-su-contenido/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GetTextLength
' Autor original    : Hans Vogelaar
' Creado            : 24/06/2012
' Fuente original   : https://social.msdn.microsoft.com/Forums/en-US/2727e4a4-57a3-4e4d-a20a-314464579ad3/
'                             how-to-calculate-the-width-of-a-access-form-textbox-pending-on-font-and-length-of-characters-string?forum=isvvba
' Propósito         : adaptar el tamaño del textbox a la longitud del texto que contiene
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Adáptalo con tus datos y pulsa F5 para ver su funcionamiento.
'
'Public Sub AutoFit(ctl As Control)
'Dim lngWidth As Long
'
'    lngWidth = GetTextLength(ctl, ctl.value)
'    ctl.Width = lngWidth + 40
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim lx As Long, ly As Long
    
' Initialize WizHook
    WizHook.key = 51488399
' Populate the variables lx and ly with the width and height of the
' string in twips, according to the font settings of the control
    WizHook.TwipsFromFont pCtrl.FontName, pCtrl.FontSize, pCtrl.FontWeight, _
                          pCtrl.FontItalic, pCtrl.FontUnderline, 0, _
                          str, 0, lx, ly
    If Not Height Then
        GetTextLength = lx
    Else
        GetTextLength = ly
    End If

End Function
