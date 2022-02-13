'Código a incorporar en el formulario
Private Sub Form_Load()
 
      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_BLEND
'      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_HOR_POSITIVE Or AW_SLIDE
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_HOR_NEGATIVE Or AW_SLIDE
'      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_CENTER
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_VER_POSITIVE Or AW_SLIDE
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_VER_NEGATIVE Or AW_SLIDE

End Sub

Private Sub Form_Unload(Cancel As Integer)

      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_BLEND Or AW_HIDE
'      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_HOR_POSITIVE Or AW_SLIDE Or AW_HIDE
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_HOR_NEGATIVE Or AW_SLIDE Or AW_HIDE
'      GeneraEfecto Me.hWnd, TiempoAnimacion, AW_CENTER Or AW_HIDE
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_VER_POSITIVE Or AW_SLIDE Or AW_HIDE
'      GeneraEfecto Me.hwnd, TiempoAnimacion, AW_VER_NEGATIVE Or AW_SLIDE Or AW_HIDE

End Sub
