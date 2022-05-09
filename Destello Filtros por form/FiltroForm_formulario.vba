'Este cloque de código es a nivel de formulario
Option Compare Database
Option Explicit

Private Sub cuepob_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = acRightButton Then
    CommandBars("Filtros de formulario").ShowPopup
    CommandBars("Filtros de formulario").Controls(2).Visible = False
End If

End Sub

Private Sub Form_Open(Cancel As Integer)

Call CrearFiltros

CommandBars("Filtros de formulario").Controls(2).Visible = False

End Sub

Private Sub nombrecuenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = acRightButton Then
    CommandBars("Filtros de formulario").ShowPopup
    CommandBars("Filtros de formulario").Controls(2).Visible = False
End If

End Sub