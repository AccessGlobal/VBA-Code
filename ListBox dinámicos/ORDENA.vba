'Función que se asocia a cada botón de comando
Private Function ORDENA()
Dim c As Integer, b As Integer
Dim NBT As String, SUF As String
Dim NCA As String

   For c = 0 To Me.l_campos.ListCount
        If Me.l_campos.Selected(c) = True Then b = b + 1: Me.Controls("BT" & b).Caption = Me.l_campos.Column(0, c)
   Next c
   
   NBT = Me.ActiveControl.Name
   NCA = Controls("" & NBT & "").Caption
   
   If Me.TX_ORDEN = "ORDER BY " & NCA & " ASC" Then SUF = ChrW(9660): Me.TX_ORDEN = "ORDER BY " & NCA & " DESC" Else Me.TX_ORDEN = "ORDER BY " & NCA & " ASC": SUF = ChrW(9650)
   
   Controls("" & NBT & "").Caption = SUF & " " & NCA & " " & SUF
   
   CARGA_SELECT
   
End Function