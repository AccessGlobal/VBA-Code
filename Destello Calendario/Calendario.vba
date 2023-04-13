'Módulo estándar: basFiestas
Option Compare Database
Option Explicit

Function GeneraResumen(Ano As String)
'--------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-sencillo-y-rapido-calendario-personalizado/
'--------------------------------------------------------------------------------------------------------------------------------
' Título            : GeneraResumen
' Autor original    : Alba Salvá
' Creado            : Desconocido
' Propósito         : Borra la tabla con el resumen de No Laborables y la vuelve a rellenar con los correspondientes al año
'                     pasado como parametro
' Argumentos        : La sintaxis de la función consta del siguiente argumento
'                     Variable     Modo             Descripción
'--------------------------------------------------------------------------------------------------------------------------------
'                     Ano          Obligatorio      Año del que queremos mostrar el calendario
'--------------------------------------------------------------------------------------------------------------------------------
Dim rs As DAO.Recordset
Dim mSql As String
Dim strNombreFiesta As String, strTipo As String
Dim mFecha As Date
Dim mDia As Byte
Dim mPaso
Dim i As Integer

    mSql = "DELETE * FROM tblResumen"
    CurrentDb.Execute mSql 'vaciamos la tabla tblResumen
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblDefineFestivo WHERE Incluir = true") 'Abrimos ahora la tabla de Festivos
    
    If Not rs.EOF And Not rs.BOF Then 'Si hay registros
       rs.MoveLast
       rs.MoveFirst
       For i = 1 To rs.RecordCount 'bucle para recorrer todos los registros
           strNombreFiesta = rs!NombreFiesta 'En este caso la descripcion es el nombre de la fiesta
           strTipo = rs!TipoFiesta 'Indica si es nacional, regional o local (población)
           mPaso = 0
           Select Case rs!FechaFija 'Segun el valor de Fecha Fija
               Case True 'Esta fiesta tiene fecha fija
                   mFecha = DateSerial(Ano, rs!Mes, rs!DiaMes) 'Se asigna directamente la fecha
               Case False 'Esta fiesta no tiene fecha fija
                   If rs!TipoFiestaVariable = 1 Then 'Si tipo fiestaVariable es un Dia de Semana y Mes (por ejemplo segundo martes de Junio)
                       mDia = rs!DiaSemana + (8 - Format("1/" & rs!Mes & "/" & Ano, "w", vbMonday)) 'Primera vez que ese dia de semana se encuentra en el mes
                       If mDia > 7 Then mDia = mDia - 7 'si sale mayor de 7 ==> le restamos 7, nos hemos ido una semana adelante
                       mDia = mDia + (7 * (rs!NumeroDiaSemana - 1)) 'Le sumamos tantas semanas como sea necesario a mDia
                       mFecha = DateSerial(Ano, rs!Mes, mDia) 'construimos la fecha completa
                   ElseIf rs!TipoFiestaVariable = 2 Then 'Fecha en funcion del Domingo de Ramos
                       mFecha = DRamos(Val(Ano)) + rs!SumarADomingoRamos 'calculamos el Domingo de Ramos y le sumamos los dias necesarios
                   End If
               End Select
               If rs!PasaLunes = True Then 'Si Pasar a Lunes está a Verdadero
                   If DatePart("w", mFecha, vbMonday) = 7 Then 'Miramos si la fecha calculada es un domingo
                       mFecha = mFecha + 1 'si lo es le sumamos un día. Se puede sustituir por 'mFecha = DateAdd("d", 1, mFecha)
                       mPaso = -1 'y ponemos a true el booleano que indica que se ha realizado el paso
                   End If
               End If
               'insertamos el registro en la tblResumen
               mSql = "INSERT INTO tblResumen (Tipo, Descripcion, TipoFiesta, Fecha, PasadoALunes) VALUES ('F', '" & strNombreFiesta & "', '" & strTipo & "', #" & Format(mFecha, "mm/dd/yyyy") & "#, " & Val(mPaso) & ")"
               CurrentDb.Execute mSql
           rs.MoveNext
       Next i
       Exit Function
    End If
End Function

Public Function DRamos(Ano As Integer) As Date
'--------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-sencillo-y-rapido-calendario-personalizado
'--------------------------------------------------------------------------------------------------------------------------------
' Título            : DRamos
' Autor original    : Alba Salvá
' Creado            : Desconocido
' Propósito         : Funcion que calcula la fecha del Domingo de Ramos para el año pasado como parámetro
' Argumentos        : La sintaxis de la función consta del siguiente argumento
'                     Variable     Modo             Descripción
'--------------------------------------------------------------------------------------------------------------------------------
' Retorno           : un valor fecha con el día en el que se celebra el "Domingo de Ramos"
'--------------------------------------------------------------------------------------------------------------------------------
Dim e As Integer, a As Integer, b As Integer, c As Integer, d As Integer
    
    a = Ano Mod 19
    b = Ano Mod 4
    c = Ano Mod 7
    d = (19 * a + 24) Mod 30
    e = (2 * b + 4 * c + 6 * d + 5) Mod 7
    
    DRamos = DateSerial(Ano, 3, 15 + d + e)

End Function

Function EsFestivo(dtFecha As Date) As Boolean

    EsFestivo = DCount("Fecha", "tblResumen", "Fecha = #" & Format(dtFecha, "mm/dd/yyyy") & "#") > 0

End Function

Function EsJI(dtDate As Date) As Boolean
Dim boJInt As Boolean
Dim dtIni As Date
Dim dtFin As Date
    
    dtIni = DFirst("dtIniJI", "tbmJI")
    dtFin = DFirst("dtFinJI", "tbmJI")
    
    EsJI = DFirst("boHayJI", "tbmJI") And _
           ( _
           dtDate >= DateSerial(Year(dtDate), Month(dtIni), Day(dtIni)) And _
           dtDate <= DateSerial(Year(dtDate), Month(dtFin), Day(dtFin)) _
           )

End Function


'Form: calendario

Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
Dim i As Integer
    
    Me.lblAñoSanto.Visible = False
    
    For i = 1 To 12
       Me.lblAño.Caption = Nz(Me.OpenArgs, Year(Date)) 'Asignamos el año pasado en el OpenArgs a la etiqueta lblYear de cada Mes
       Me.Controls(i & "lblYear").Caption = Format(Nz(Me.OpenArgs, Year(Date)), "#,##0")
       Me.Controls(i & "lblMes").Caption = UCase(Format("1/" & i, "mmmm")) 'Asignamos el nombre del mes a cada Mes
       Me.Controls(i & "lblMesEnNumero").Caption = i 'Asignamos el numero del mes a cada Mes
    Next i
   
   LlenaDias 'Llamamos al procedimiento que rellenará el calendario

   If Weekday(DateSerial(Me.lblAño.Caption, 7, 25)) = vbSunday Then
       Me.lblAñoSanto.Visible = True
   End If
    
End Sub

Sub LlenaDias()
'--------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-sencillo-y-rapido-calendario-personalizado
'--------------------------------------------------------------------------------------------------------------------------------
' Título            : LlenaDias
' Autor original    : Desconocido
' Creado            : Alba Salvá
' Propósito         : completa el calendario atendiendo a los festivos del año calculado
'--------------------------------------------------------------------------------------------------------------------------------
Dim PrimerDia As Date, UltimoDia As Date, DiaActual As Date, PrimerDomingo As Date
Dim PrimerDiaSemana, DiaTemporal As Long
Dim frmActual As Form
Dim ctrActual As control
Dim xCont As Long
Dim Nomcc As String
Dim ccForm As control
Dim strTipoNL As String
Dim poscc
Dim m As Integer
Dim fActual As Date
Dim cL, cT
   
   Set frmActual = Me
   
    If Year(DFirst("fecha", "tblresumen")) <> Me.lblAño.Caption _
    Or IsNull(Year(DFirst("fecha", "tblresumen"))) Then
        GeneraResumen Me.lblAño.Caption
    End If

Me.cmdEli.Visible = False
For m = 1 To 12 'bucle que se repite para cada mes
    DiaTemporal = 1
    'En PrimerDia establecemos las fecha completa (DateSerial) del primer dia del mes
    PrimerDia = DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, 1)
    'En UltimoDia la del ultimo
    UltimoDia = DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption + 1, 0)
    'En PrimerDiaSemana determinamos que dia de semana es PrimerDia
    PrimerDiaSemana = Weekday(PrimerDia, vbUseSystemDayOfWeek)
    'En PrimerDomingo determinamos el numero de dias de la primera semana
    PrimerDomingo = PrimerDia - Weekday(PrimerDia, vbUseSystemDayOfWeek) + 7
    xCont = 0
   
    For Each ctrActual In frmActual 'Recorremos todos los controles de formulario
     poscc = InStr(1, ctrActual.Name, "cc") - 1 'En poscc guardamos la longitud el numero inicial en el nombre del control (1 o 2 cifras) que identifica el mes
     If poscc < 0 Then poscc = 0 'Hay controles (los que no son para mostrar los dias) que no cumplen la condicion. Así evitamos el error
         If Mid(ctrActual.Name, 1, poscc) = m Then 'si el control pertenece al mes que estamos rellenando
             Set ccForm = ctrActual
             Nomcc = m & "cc" & xCont 'construimos el nombre del campo que nos correspondería rellenar
             xCont = xCont + 1
             
             If ccForm.Name = Nomcc Then 'si el nombre del control actual es el que nos correspondería rellenar
                
                If xCont >= PrimerDiaSemana Then 'si ya hemos sobrepasado los dias vacios a principio de mes
                   ccForm.Caption = DiaTemporal 'asignamos al control el numero del dia correspondiente
                   DiaTemporal = DiaTemporal + 1 'incrementamos en 1 el numero de dia a rellenar
                   ccForm.Visible = True 'hacemos visible el control
                   '
                   ' TRATAMIENTO PARA LOS NO LABORABLES
                   
                   fActual = DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1)
                   
                   strTipoNL = Nz(DLookup("Tipo", "tblResumen", "Fecha = #" & Format(DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1), "mm/dd/yyyy") & "#"), "")
                      
                   If strTipoNL = "F" Then 'si es un FESTIVO
                      ccForm.ForeColor = 16777215 'blanco: color de fuente
                      ccForm.BackColor = 16711680 'azul: color de fondo
                      ccForm.FontWeight = 400 'tamaño de fuente
                      'Asignamos a la propiedad "Texto de ayuda al Control" el texto explicativo del festivo
                      ccForm.ControlTipText = DLookup("Descripcion", "tblResumen", "Fecha = #" & Format(DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1), "mm/dd/yyyy") & "#") & vbCrLf & _
                                        "(" & DLookup("TipoFiesta", "tblResumen", "Fecha = #" & Format(DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1), "mm/dd/yyyy") & "#") & ")"
'Añadido por Alba Salvá
'*************************
                   ElseIf Weekday(fActual) = vbSaturday Then 'Si es Sábado
                      ccForm.ForeColor = vbBlack ' negro: color de fuente
                      ccForm.BackColor = vbYellow 'amarillo: color de fondo
                      ccForm.FontWeight = 400 'tamaño de fuente
                      'Asignamos a la propiedad "Texto de ayuda al Control" el texto explicativo del festivo
                      ccForm.ControlTipText = "Sábado"
    
                   ElseIf Weekday(fActual) = vbSunday Then 'Si es Domingo
                      ccForm.ForeColor = 16777215 'blanco color de fondo
                      ccForm.BackColor = 255 'rojo: color de fuente
                      ccForm.FontWeight = 400 'tamaño de fuente
                      'Asignamos a la propiedad "Texto de ayuda al Control" el texto explicativo del festivo
                      ccForm.ControlTipText = "Domingo"
    
                   ElseIf EsJI(fActual) Then 'Si es Jornada intensiva
                     ccForm.ForeColor = 16711680 'azul: color de fuente
                     ccForm.BackColor = 5167783 'verde claro: color de fondo
                     
'*************************
                   Else 'si es LABORABLE
                      ccForm.ForeColor = 0 'negro: color de fuente
                      ccForm.FontWeight = 400
                      ccForm.BackColor = 16777215 'blanco color de fondo
                      ccForm.ControlTipText = "" 'vaciamos la propiedad Texto de ayuda al Control
                   End If
                         
                   If DiaTemporal > Day(UltimoDia) + 1 Then 'Si este era el ultimo dia del mes
                      ccForm.Caption = ""
                      ccForm.Visible = False
                   End If
                
'Esto es mio.
'*************************
                   If fActual = Date Then
                      cL = Me.cmdEli.Width / 2 - ccForm.Width / 2
                      cT = Me.cmdEli.height / 2 - ccForm.height / 2
                      
                      Me.cmdEli.Left = ccForm.Left - cL
                      Me.cmdEli.Top = ccForm.Top - cT
                      Me.cmdEli.Visible = True
                      
                      If EsFestivo(Date) Then
                        Me.cmdEli.ControlTipText = DLookup("Descripcion", "tblResumen", "Fecha = #" & Format(DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1), "mm/dd/yyyy") & "#") & vbCrLf & _
                                             "(" & DLookup("TipoFiesta", "tblResumen", "Fecha = #" & Format(DateSerial(Me.Controls(m & "lblYear").Caption, Me.Controls(m & "lblMesEnNumero").Caption, DiaTemporal - 1), "mm/dd/yyyy") & "#") & ")"
                      End If
                   End If
'*************************
                Else
                   DiaTemporal = 1
                   ccForm.Caption = ""
                   ccForm.Visible = False
                End If
             End If
         End If
    Next ctrActual
Next m
End Sub

Private Sub btnCerrar_Click()
On Error GoTo Err_btnCerrar_Click


    DoCmd.Close acForm, Me.Name, acSaveNo 'cerrar formulario

Exit_btnCerrar_Click:
    Exit Sub

Err_btnCerrar_Click:
    MsgBox Err.Description
    Resume Exit_btnCerrar_Click
    
End Sub

Private Sub btnImprimir_Click()
On Error GoTo Err_btnImprimir_Click

    DoCmd.PrintOut 'imprimir formulario

Exit_btnImprimir_Click:
    Exit Sub

Err_btnImprimir_Click:
    MsgBox Err.Description
    Resume Exit_btnImprimir_Click

End Sub

'Forma: DefinirFestivos
Option Compare Database
Option Explicit

Private Sub btnCerrar_Click()

    DoCmd.Close acForm, Me.Name, acSaveNo
    
End Sub

Private Sub cboTipoFiestaVariable_Click()
    Me.chkFechaFija = Me.cboTipoFiestaVariable = 0
    
    Me.cboDiaSemana.Enabled = False
    Me.txtNumeroDiaSemana.Enabled = False
    
    Me.txtDiaMes.Enabled = False
    Me.cboMes.Enabled = False
    
    Me.txtSumarADomingoRamos.Enabled = False
    
    Select Case Me.cboTipoFiestaVariable
        Case 0
            Me.txtDiaMes.Enabled = True
            Me.cboMes.Enabled = True
        Case 1
            Me.cboDiaSemana.Enabled = True
            Me.txtNumeroDiaSemana.Enabled = True
            Me.cboMes.Enabled = True
        Case 2
            Me.txtSumarADomingoRamos.Enabled = True
        Case Else
    End Select
    
End Sub

Private Sub Form_Close()

    GeneraResumen CStr(Year(Date))
    
End Sub

Private Sub Form_Current()

    Me.chkFechaFija = Me.cboTipoFiestaVariable = 0
    
    Me.cboDiaSemana.Enabled = False
    Me.txtNumeroDiaSemana.Enabled = False
    
    Me.txtDiaMes.Enabled = False
    Me.cboMes.Enabled = False
    
    Me.txtSumarADomingoRamos.Enabled = False
    
    Select Case Me.cboTipoFiestaVariable
        Case 0
            Me.txtDiaMes.Enabled = True
            Me.cboMes.Enabled = True
        Case 1
            Me.cboDiaSemana.Enabled = True
            Me.txtNumeroDiaSemana.Enabled = True
            Me.cboMes.Enabled = True
        Case 2
            Me.txtSumarADomingoRamos.Enabled = True
        Case Else
    End Select
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim strTemp As String, i As Integer


    Me.lblAño.Caption = Nz(Me.OpenArgs, Year(Date)) 'Asignamos el año pasado en el OpenArgs a la etiqueta lblYear
    
    With Me.cboTipoFiestaVariable
        .RowSourceType = "Value List"
        .RowSource = ""
        .RowSource = "0;'Es fecha fija';1;'Es un Dia de Semana y Mes';2;'Depende del Domingo de Ramos'"
        .Requery
    End With
    
    strTemp = "0;''"
    For i = 1 To 7
        strTemp = strTemp & ";" & i & ";'" & StrConv(Format(DateSerial(2000, i, 1), "dddd"), vbProperCase) & "'"
    Next
    With Me.cboDiaSemana
        .RowSourceType = "Value List"
        .RowSource = ""
        .RowSource = strTemp
        .Requery
    End With
    
    strTemp = "0;''"
    For i = 1 To 12
        strTemp = strTemp & ";" & i & ";'" & StrConv(Format(DateSerial(2000, i, 1), "mmmm"), vbProperCase) & "'"
    Next
    With Me.cboMes
        .RowSourceType = "Value List"
        .RowSource = ""
        .RowSource = strTemp
        .Requery
    End With
    
End Sub


