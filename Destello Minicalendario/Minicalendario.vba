'Poner este código en el form del minicalendario
' Formulario Calendario_mini
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-minicalendario/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Calendario_mini
' Autor original    : Antonio Otero
' Fecha             : febrero 22
' Propósito         : disponer de todos los elementos y funciones necesarias para la creación de un minicalendario
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub CB_ANO_AfterUpdate()
    
    OBTENER_DIAS

End Sub

Function OBTENER_DIAS()
    Dim an As Integer
    Dim mesv As Integer
    Dim primerdia As Date
    Dim lngdiasmes As String
    Dim lngprimerdia As String
    Dim m As Integer
    Dim dini As Integer
    Dim dfin As Integer
    Dim d As Integer
    Dim ff As Date
    Dim dias As Integer
    Dim filb As String
    Dim ex As Variant
    
    an = Me.CB_ANO
    mesv = Me.ETNMES.Caption
    
     primerdia = DateSerial(an, mesv, 1)
     lngdiasmes = Day(DateSerial(an, mesv + 1, 1) - 1)
     lngprimerdia = DatePart("w", primerdia, vbMonday)

   
    For m = 1 To 42
            Me.Controls("D" & m).Visible = False
            Me.Controls("d" & m).BackColor = vbWhite
            Me.Controls("d" & m).Value = 0
    Next m
    dini = lngprimerdia: dfin = dini + (lngdiasmes - 1)
    
    d = 0
    For m = dini To dfin
            d = d + 1
            Me.Controls("D" & m).Visible = True
            Me.Controls("d" & m).Caption = d
            ff = Format(Format(d, "00") & "/" & Format(mesv, "00") & "/" & an, "mm/dd/yyyy")
            dias = Weekday(ff)
            If dias = 1 Then Me.Controls("d" & m).BackColor = vbBlue
            ex = DLookup("fecha", "public_festivos", "fecha = #" & ff & "#")
            If Not IsNull(ex) Then Me.Controls("d" & m).BackColor = vbRed
    Next m
            
End Function

Private Sub cb_mes_AfterUpdate()
    Dim mex  As Variant
    
    
    Dim n As Integer
    
    mex = Array("", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    For n = 1 To 12
        If Me.cb_mes = mex(n) Then Me.ETNMES.Caption = n
    Next n
    OBTENER_DIAS
    
End Sub

Private Function pulsarb()
        Dim ncon As String, fpul As String
        Dim colo As Double, ncolo As Double
        Dim diap As Integer
        Dim ff As Date
        
        
         ncon = Me.ActiveControl.Name
         colo = Me.Controls("" & ncon & "").BackColor
         diap = Me.Controls("" & ncon & "").Caption
         If colo = 16711680 Or colo = 255 Then ncolo = 16711680: Me.Controls("" & ncon & "").Value = 0
        
        fpul = Format(diap, "00") & "/"
        ff = Format(diap, "00") & "/" & Format(Me.ETNMES.Caption, "00") & "/" & Me.CB_ANO
        Me.tx_fecha = Format(ff, "dddd, dd  mmmm , yyyy")
        If colo = 16777215 Then colo = vbBlack
        Me.tx_fecha.ForeColor = colo
End Function

Private Sub D7_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
    If IsNull(Me.CB_ANO) Then Me.CB_ANO = Year(Date)
    If IsNull(Me.cb_mes) Then Me.cb_mes = Format(Date, "mmmm")
    cb_mes_AfterUpdate
        
End Sub