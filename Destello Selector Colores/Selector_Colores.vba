'Formulario MisColores
Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-mi-selector-de-colores/
'					  Destello formativo 343
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : frm_MisColores
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Fecha             : en algún momento del verano de 2012
' Propósito         : disponer de un selector de colores personalizado
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Declare PtrSafe Sub wlib_AccColorDialog Lib "msaccess.exe" Alias "#53" (ByVal hwnd As Long, lngRGB As Long)

Private Sub btnCerrar_Click()

    DoCmd.Close acForm, "MisColores"

End Sub

Private Sub btnP01_Click()
    
    Me.P01.BackColor = SelecColor(Me.P01.BackColor)

End Sub

Private Sub btnP02_Click()
    
    Me.P02.BackColor = SelecColor(Me.P02.BackColor)

End Sub

Private Sub btnP03_Click()
    
    Me.P03.BackColor = SelecColor(Me.P03.BackColor)

End Sub

Private Sub btnP04_Click()
    
    Me.P04.BackColor = SelecColor(Me.P04.BackColor)

End Sub

Private Sub btnP05_Click()
    
    Me.P05.BackColor = SelecColor(Me.P05.BackColor)

End Sub

Private Sub btnP06_Click()
    
    Me.P06.BackColor = SelecColor(Me.P06.BackColor)

End Sub

Private Sub btnP07_Click()
    
    Me.P07.BackColor = SelecColor(Me.P07.BackColor)

End Sub

Private Sub btnP08_Click()
    
    Me.P08.BackColor = SelecColor(Me.P08.BackColor)

End Sub

Private Sub btnP09_Click()
    
    Me.P09.BackColor = SelecColor(Me.P09.BackColor)

End Sub

Private Sub btnP10_Click()
    
    Me.P10.BackColor = SelecColor(Me.P10.BackColor)
    
End Sub

Private Sub Form_Close()
Dim col As Long
Dim i As Integer
Dim P As String, colhex As String, colrgb As String
Dim collng As Long
Dim cont As control
Dim R, G, B
Dim rstTable As DAO.Recordset

    For Each cont In Me.Controls
        P = Left(cont.Name, 1)
        If P = "P" Then
            If Left(cont.Name, 2) = "P0" Then
                i = Right(cont.Name, 2)
                col = Abs(cont.BackColor)
            Else
                col = Abs(cont.BackColor)
                i = Right(cont.Name, 2)
            End If

'Color lng
            collng = CLng("&H" & Right("000000" + _
                     Replace(Nz(cont.BackColor, ""), "#", ""), 6))
    
'Color RGB
            R = col Mod 256
            G = (col \ 256) Mod 256
            B = (col \ 256 \ 256) Mod 256
            
            colrgb = "(" & R & "," & G & "," & B & ")"
            
'Color hex
            R = hex(R)
            G = hex(G)
            B = hex(B)
                
            If Len(B) = 1 Then B = 0 & B
            colhex = "#" & R & G & B
            Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM colores WHERE idcolores=" & i)
                rstTable.Edit
                    rstTable!colorint = cont.BackColor
                    rstTable!colorlng = collng
                    rstTable!colorrgb = colrgb
                    rstTable!colorhex = colhex
                rstTable.Update
                rstTable.Close
            Set rstTable = Nothing
        End If
    Next
       
End Sub

Private Sub Form_Open(cancel As Integer)
Dim i As Integer
Dim cont As control
Dim lit1 As String
Dim rstTable As DAO.Recordset

    lit1 = "Haga clic en el color para cambiar el color"

    Me.P01.ControlTipText = lit1
    Me.P02.ControlTipText = lit1
    Me.P03.ControlTipText = lit1
    Me.P04.ControlTipText = lit1
    Me.P05.ControlTipText = lit1
    Me.P06.ControlTipText = lit1
    Me.P07.ControlTipText = lit1
    Me.P08.ControlTipText = lit1
    Me.P09.ControlTipText = lit1
    Me.P10.ControlTipText = lit1

    Set rstTable = CurrentDb.OpenRecordset("colores")
        Do Until rstTable.EOF
            i = rstTable!IdColores
            If IsNull(rstTable!colorrgb) Then GoTo LinNext
            
            For Each cont In Me.Controls
                If cont.Name = "P0" & i Then cont.BackColor = rstTable!colorint: Exit For
                If cont.Name = "P" & i Then cont.BackColor = rstTable!colorint: Exit For
            Next
LinNext:
        rstTable.MoveNext
        Loop

    Set rstTable = Nothing

End Sub

Private Function SelecColor(MiColor As Variant) As Long
Dim lngColor As Long

    lngColor = CLng("&H" & Right("000000" + _
                       Replace(Nz(MiColor, ""), "#", ""), 6))
    
    wlib_AccColorDialog Screen.ActiveForm.hwnd, lngColor
    
    SelecColor = lngColor

End Function