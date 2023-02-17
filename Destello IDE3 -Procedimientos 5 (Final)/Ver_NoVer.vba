'Diferentes eventos de un formulario
Private Sub btnNoVer_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-consideraciones-finales-y-una-funcion-inutil
'                     Destello formativo 268
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Ver/No ver
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : Muestra y oculta el IDE y captura su título
'---------------------------------------------------------------------------------------------------------------------------------------------------
    Application.VBE.MainWindow.Visible = False
    
End Sub

Private Sub btnVer_Click()

     Application.VBE.MainWindow.Visible = True
     
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Me.TituloTxt.Caption = Application.VBE.MainWindow.Caption
    
End Sub