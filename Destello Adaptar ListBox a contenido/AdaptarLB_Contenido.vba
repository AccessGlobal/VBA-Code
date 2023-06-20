'Evento Al abrir del formulario que contiene el listbox
Private Sub Form_Open(Cancel As Integer)
Dim strSQL As String
Dim rstTable As DAO.Recordset
Dim strLng As String
Dim i As Integer
Dim matLen()
Dim wzFontName As String
Dim wzSize As Long
Dim wzWeight As Long
Dim wzItalic As Boolean
Dim wzUnderline As Boolean
Dim wzCch As Long
Dim wzCaption As String
Dim wzMaxWidthCch As Long
Dim wzdx As Long
Dim wzdy As Long
Dim lstBox As ListBox

    Set lstBox = Me.LstProductos
    
        strSQL = "SELECT idprodut, produtcod, produtnom, precio, activo FROM productos;"
           
'Para la lingitud de las columnas calculamos los twips que tiene un caracter
        WizHook.key = 51488399
        wzFontName = lstBox.FontName
        wzSize = lstBox.FontSize
        wzWeight = lstBox.FontWeight
        wzItalic = lstBox.FontItalic
        wzUnderline = lstBox.FontUnderline
        wzCaption = " "
        
        WizHook.TwipsFromFont wzFontName, wzSize, wzWeight, _
                              wzItalic, wzUnderline, wzCch, _
                              wzCaption, wzMaxWidthCch, _
                              wzdx, wzdy
'Construímos la matriz que contenga las dimensiones de las columnas autoajustadas
        matLen = AutoAjusteListbox(strSQL, wzdx)
            strLng = "0cm;"
        For i = 1 To UBound(matLen) - 1
            strLng = strLng & matLen(i) & "cm;"
        Next i
'Construímos el listbox
        Me.LstProductos.ColumnCount = 5
        Me.LstProductos.ColumnWidths = strLng
        Me.LstProductos.RowSource = strSQL
    
    Set lstBox = Nothing
    
End Sub

'Módulo estándar
Option Compare Database
Option Explicit

Public Const twipscm = 566.9291338583

Public Function AutoAjusteListbox(strSQL As String, wzdx As Long) As Variant
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-adaptar-columnas-de-un-listbox-a-su-contenido/
'                     Destello formativo 342
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AutoAjusteListbox
' Autor             : Luis Viadel | luisviadel@access-global.net
' Creado            : 16/06/2023
' Propósito         : ajustar las columnas de un listbox al contenido del mismo
' Más información   : https://access-global.net/vba-alinear-campo-de-listbox-a-la-derecha/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Adáptalo con tus datos y pulsa F5 para ver su funcionamiento.
'
'Public Sub Micampo_AfterUpdate()
'Dim matAncho()
'
'    matAncho=AutoAjusteListbox(strSQL, wzdx)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim rstTable As DAO.Recordset, rstTable1 As DAO.Recordset
Dim i As Integer
Dim matcampos()
Dim matLen()
Dim dblCampo As Double, dblCampoTotal As Double

'Ponemos control de errores para valores nulos
    On Error Resume Next
    
'Ponemos los campos de la consulta pasada como parámetro en una matriz
    Set rstTable = CurrentDb.OpenRecordset(strSQL)
        ReDim matcampos(rstTable.Fields.Count) 'Matriz con el nombre de los campos
        ReDim matLen(rstTable.Fields.Count) 'Matriz con la longitud del campo más largo
        
        For i = 1 To rstTable.Fields.Count - 1
            matcampos(i) = rstTable.Fields(i).name
'Recorremos cada campo para buscar el campo más largo
                Set rstTable1 = CurrentDb.OpenRecordset(strSQL)
                    Do Until rstTable1.EOF
'Tenemos la longitud en número de caracteres, pasamos a twips mediante WizHook (Destello 341)
                        dblCampo = Len(rstTable1.Fields(i).value) * 2 * wzdx 'Medido en twips
                        dblCampo = dblCampo / twipscm 'Pasamos a centímetros
                        dblCampo = FormatNumber(dblCampo, 2)
                        If dblCampoTotal < dblCampo Then dblCampoTotal = dblCampo
                    rstTable1.MoveNext
                    Loop
                Set rstTable1 = Nothing
            matLen(i) = dblCampoTotal + 0.4 'Ajuste tamaño barras y espacios anteriores y posteriores
'Reinicializamos para el siguiente campo
            dblCampo = 0
            dblCampoTotal = 0
        Next i
    Set rstTable = Nothing
    
    AutoAjusteListbox = matLen

End Function