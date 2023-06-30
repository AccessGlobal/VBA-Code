Private Sub MiListBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-Seleccionar-un-dato-concreto-de-un-listbox/
'                     Destello formativo 349
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : MiListBox_MouseMove
' Autor original    : Antonio Otero | antoniootereo@access-global.net
' Creado            : desconocido
' Propósito         : obtener la coordenada X del ratón para localizar la fila y la columna en un listbox
' Argumentos        : los argumentos del evento "Al mover el mouse"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ancho As String, wcol As Variant, k As Integer, tanc As Double
Dim NCol As Integer
'Dim tx as string

    ancho = Me.MiListBox.ColumnWidths
    wcol = Split(ancho, ";")

    For k = 0 To UBound(wcol)

        tanc = tanc + wcol(k)

        If k = 0 Then
            NCol = 1
        Else
'            tx = "X:" & X & vbCrLf & "tanc:" & tanc & vbCrLf & "wcol(K-1):" & wcol(k - 1) & vbCrLf & " x<= tanc and x<= wcol(k-1)" & vbCrLf & ancho
            If X <= tanc And X >= wcol(k - 1) Then
                NCol = k + 1
                Exit For
            End If
        End If

    Next k
    
    Me.NCol = NCol

    Me.Valor = Me.MiListBox.column(NCol - 1)
  
End Sub
