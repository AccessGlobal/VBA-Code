Private Sub l_campos_Click()
    Dim lar()
    Dim izqi As String, izq As String
    Dim lanch As String
    Dim i%, e%, b%
    
        lar = Array(500, 4000, 4000, 2000, 2500, 1000)
        
        izqi = Me.Lista4.Left
        izq = izqi
        For i = 1 To 6: Me.Controls("bt" & i).Visible = False: Next i
        With Me.l_campos
             
                For e = 0 To .ListCount - 1
                   If .Selected(e) = True Then
                        b = b + 1
                        Me.Controls("bt" & b).Visible = True
                        Me.Controls("bt" & b).Width = lar(e)
                        Me.Controls("bt" & b).Left = izq
                        Me.Controls("bt" & b).Caption = .Column(0, e)
                        izq = izq + lar(e) + 100
                        
                        lanch = lanch & Format((lar(e) + 100) * 0.001763, "0.00") & " cm;"
                       
                   End If
                
                Next e
               If lanch <> "" Then lanch = Left(lanch, Len(lanch) - 1)
             
                
                'MsgBox lanch
      End With
         If lanch <> "" Then
            With Me.Lista4
                    .ColumnWidths = lanch
                    .ColumnCount = b
                    .Width = izq - izqi
                 
            End With
         End If
         
         CARGA_SELECT
         
    End Sub
    
    Function CARGA_SELECT()
    Dim c As Integer
    Dim SQ As String
    
        For c = 0 To Me.l_campos.ListCount
            If Me.l_campos.Selected(c) = True Then SQ = SQ & Me.l_campos.Column(0, c) & ","
        Next c
        
        SQ = " SELECT " & Left(SQ, Len(SQ) - 1) & " FROM CLIENTES " & Me.TX_ORDEN
       
        Me.Lista4.RowSource = SQ
        
    End Function