'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-auditoria/
'                     Destello formativo 348
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Auditoría
' Autor             : Luis Viadel | luisviadel@access-global.net
' Creado            : 28/06/2023
' Propósito         : sistema para registrar los cambios en los datos de nuestra aplicación
' Más información   : 
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test              : 
'-----------------------------------------------------------------------------------------------------------------------------------------------


'Código para incorporar en el formulario del campo que deseamos controlar
Private Sub Form_Open(cancel As Integer)
    
     Call Auditoria(0, "", Me.Name, "Ha accedido a " & Me.Name)

End Sub

Private Sub Form_Close()

'Necesario para auditoría
    TempVars.RemoveAll
    
End Sub

Private Sub MiCampo_GotFocus()

'Necesario para auditoría
    TempVars!ValorInicial = Me.MiCampo.Value
    
End Sub

Private Sub MiCampo_LostFocus()
Dim audittxtExt As String
   
'Comprobación para auditoría
    If TempVars!ValorInicial <> Me.MiCampo Then
        udittxtExt = "El campo ha cambiado de " & TempVars!ValorInicial & " a " & Me.MiCampo 
        Call Auditoria(Me.MiId, "NombreCampo", Me.Name, "Acción realizada", audittxtExt)
    End If
    
End Sub

'Código en módulo estándar
Option Compare Database
Option Explicit

Public Function Auditoria(idaudit As Integer, campotxt As String, formtxt As String, AUDITTxt As String, Optional audittxtext As String)
Dim strDate As String, strTime As String, strcod As String
Dim rstTable As DAO.Recordset

'Creamos un código
    strDate = Format(Date, "ddmmyy")
    strTime = Format(Time, "mmss")
    strcod = "F" & strDate & strTime & "L"
    
    Set rstTable = CurrentDb.OpenRecordset("audit")
        rstTable.AddNew
            rstTable!auditf= Format(Date, "Short date")
            rstTable!audith= Format(Time, "Short time")
            rstTable!auditcampo = campotxt
            rstTable!auditcod = strcod
            rstTable!auditform = formtxt
            rstTable!audittxt= audittxt
            rstTable!audittxtext = audittxtext
'Se supone que la aplicación contiene un mecanismo para conocer el usuario actual
'Si es así, se pondría en idtraba la función que lo indique
            rstTable!idtraba = 1 ' Poner función
        rstTable.Update
        rstTable.Close
    Set rstTable = Nothing

End Function
