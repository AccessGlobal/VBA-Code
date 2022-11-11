
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-cuenta-atras/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Cuenta_Atras
' Autor             : Luis Viadel
' Fecha             : nov 2022
' Propósito         : Cómo crear una cuenta atrás aprovechando el objeto webBrowser
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación, copia los bloques siguientes al portapapeles y pega en el editor de VBA de los
'                    formularios correspondientes.
'
'-----------------------------------------------------------------------------------------------------------------------------------------------

'Form1
Private ContadorTiempo As Date
Private Const SegundosDia As Long = 86400

Private Sub Form_Load()
Const ContadorSegundos  As Long = 10 'Especifico el valor de la cuenta atrás
    
    Me.Cerrar.SetFocus 'Quito el foco del cuadro de texto para evitar parpadeos
    
    ContadorTiempo = DateAdd("s", ContadorSegundos, Now)
    Me.Contador.Value = ContadorTiempo
    
    Me.TimerInterval = 100

End Sub

Private Sub Form_Timer()
Dim TiempoRestante As Date
    
    Me.Cerrar.SetFocus 

    TiempoRestante = CDate(ContadorTiempo - Date - Timer / SegundosDia)
    
    Me.Contador.Value = TiempoRestante
    
    If TiempoRestante <= 0 Then
        Me.TimerInterval = 0
        DoCmd.Close acForm, Me.Name
        DoCmd.OpenForm "Form2"
    End If

End Sub

'-----------------------------------------------------------------------------
'Form2
Private Sub Form_Open(Cancel As Integer)
Dim ruta As String
'El GIF lo convierto en HTML
    ruta = "C:\ruta de mi fichero HTML"
    
    Me.Navegador.navigate ruta
    
End Sub

'-----------------------------------------------------------------------------
'Estructura del fichero HTML
<html>
   <head>
      <title>Cuenta atrás</title>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
   </head>
   <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
      <img src="Ruta de mi GIF" width="690" height="729" alt="">
   </body>
</html>