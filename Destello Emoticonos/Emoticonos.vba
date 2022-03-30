Option Compare Database
Option Explicit

'Código del form "frmEmoji"

Private Sub Form_Open(Cancel As Integer)
Dim rstTable1 As DAO.Recordset
Dim i As Integer, j As Integer
Dim ctrl As Control

j = 1

For i = 131 To 150
    Set rstTable1 = CurrentDb.OpenRecordset("SELECT * FROM emoticonos WHERE idemoti=" & i)
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is TextBox Then
                If j < 10 Then
                    If ctrl.Name = "e0" & j Then ctrl.Value = "<p>" & rstTable1!htmlcode & "</p>"
                Else
                    If ctrl.Name = "e" & j Then ctrl.Value = "<p>" & rstTable1!htmlcode & "</p>"
                End If
            End If
        Next ctrl
        rstTable1.Close
    Set rstTable1 = Nothing
    j = j + 1
Next i

End Sub

Private Function ponericono()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-utiliza-emoticonos-en-tus-aplicaciones
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ponericono
' Autor             : Luis Viadel | https://cowtechnologies.net
' Actualizado       : 28/03/2022
' Propósito         : al seleccionar un icono del formulario, este se escribe en el formulario de mensajes, tomando el texto que hay en el texbox
'                     y añadiéndolo el icono, según el código HTML del mismo.
'-----------------------------------------------------------------------------------------------------------------------------------------------

Dim ctrl As Control
Dim str1 As String, str2 As String

Set ctrl = Screen.ActiveControl
str1 = Right(ctrl.Name, 2)

str2 = "e" & str1

For Each ctrl In Form_frmEmoji.Controls
    If ctrl.Name = str2 Then
        str1 = Right(ctrl.Value, Len(ctrl.Value) - 3)
        str1 = Left(str1, Len(str1) - 4)
        str2 = Form_frmEmojiEscribir.txtmensaje.Value
        str2 = Left(str2, Len(str2) - 6)
        
        Debug.Print str2 & str1 & "</div>"
        If IsNull(Form_frmEmojiEscribir.txtmensaje.Value) Then
            Form_frmEmojiEscribir.txtmensaje.Value = "<div>" & str1 & "</div>"
        Else
            Form_frmEmojiEscribir.txtmensaje.Value = str2 & str1 & "</div>"
        End If
        Form_frmEmojiEscribir.Refresh
        DoCmd.Close acForm, "frmEmoji"
        Exit Function
    End If
Next
    
End Function
