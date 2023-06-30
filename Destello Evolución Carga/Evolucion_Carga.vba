'Módulo estándar de un formulario
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-evolucion-de-carga-artesano/
'                      Destello formativo 350
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ContadorRegistros
' Autor original    : Antonio Otero | antoniootero@access-global.net
' Creado            : desconocido
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Propósito         : mostrar la evolución de carga de registros de un recordset. Dos posibilidades empezando cada vez desde el principio
'                     o continuando después de una parada
'---------------------------------------------------------------------------------------------------------------------------------------------
' Objetos           : se precisan dos botones (btnCargar y btnCargar2), dos botones de stop (btnStop y btnStop2)
'                     y dos etiquetas (txtMonitor y txtMonitor2)
'---------------------------------------------------------------------------------------------------------------------------------------------

Private X%, X2%, TT%, i%

Private Sub btnCargar_Click()
Dim RS As DAO.Recordset
    
    Set RS = CurrentDb.OpenRecordset("select * from productos")
        RS.MoveLast
        TT = RS.RecordCount
        RS.MoveFirst
        X = 0
        While Not RS.EOF
            X = X + 1
            Me.txtMonitor.Caption = "Cargando registros" & vbCrLf & X & "/" & TT
            If TempVars!parar = -1 Then GoTo LinExit
            DoEvents
            RS.MoveNext
        Wend
        
    Set RS = Nothing
    
LinExit:
    TempVars.Remove ("parar")
    
End Sub

Private Sub btnCargar2_Click()
Dim RS2 As DAO.Recordset
    
    Set RS2 = CurrentDb.OpenRecordset("select * from productos")
        RS2.MoveLast
        TT = RS2.RecordCount
        RS2.MoveFirst
        If X2 <> 0 Then
            For i = X2 To TT
                Me.txtMonitor2.Caption = "Cargando registros" & vbCrLf & i & "/" & TT
                If TempVars!parar2 = -1 Then
                    X2 = i
                    TempVars!reg2 = X2
                    GoTo LinExit
                End If
                DoEvents
            Next i
        Else
            X = 0
            While Not RS2.EOF
                X2 = X2 + 1
                Me.txtMonitor2.Caption = "Cargando registros" & vbCrLf & X2 & "/" & TT
                If TempVars!parar2 = -1 Then
                    TempVars!reg2 = X2
                    GoTo LinExit
                End If
                DoEvents
                RS2.MoveNext
            Wend
        End If
    Set RS2 = Nothing
    
LinExit:
    TempVars.Remove ("parar2")
    
End Sub

Private Sub btnStop_Click()

    TempVars!parar = -1
    
End Sub


Private Sub btnStop2_Click()

    TempVars!parar2 = -1
    
End Sub