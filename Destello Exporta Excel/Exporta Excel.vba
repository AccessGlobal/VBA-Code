Sub ExportaExcel(ByVal strSQL As String, strFilename As String, ByVal pasos As Integer, Optional strSheetName As String = "", Optional boShExcel As Boolean)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-exportar-contenido-de-un-recordset-a-excel-copyfromrecordset
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExportaExcel
' Autor             : Alba Salvá
' Fecha             : no se acuerda, pero hace mucho tiempo
' Propósito         : Copia el contenido de un objeto Recordset ADO o DAO en una hoja de Excel
' Retorno           : Sin retorno
' Argumento/s       : La sintaxis del procedimiento consta del siguiente argumento:
'                     Parte            Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strSQL        Obligatorio    Datos que vamos a exportar a Excel
'                     strFilename   Obligatorio    Nombre del fichero de destino de los datos
'                     pasos         Obligatorio    Indica la cantidad de registros a insertar de cada vez,
'                                                  aumentar o disminuir en función de la velocidad de la red.
'                     strSheetName  Opcional       Nombre de la hoja
'                     boShExcel     Opcional       Activar/desactivar propiedades de la hoja que se abre
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://docs.microsoft.com/en-us/office/vba/api/excel.range.copyfromrecordset
' Importante        : Copia el contenido de un objeto Recordset ADO o DAO en una hoja de Excel, comenzando en la esquina superior izquierda
'                     del rango especificado. Si el objeto Recordset contiene campos con objetos OLE y campos multivalor, este método falla.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Copiar el bloque siguiente al portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese
'
'Sub ExportaExcel_test()
'Dim strSQL As String, StrRuta As String
'Dim pasos As Integer
'
'strSQL = "SELECT tabla.campo1, tabla.campo2, tabla.campo3, tabla.campo4, tabla.campo5 " & _
'         "FROM tabla;"
'
'StrRuta = Application.CurrentProject.Path & "\Exportar_test.xlsx"
'
'Call ExportaExcel(strSQL, StrRuta, 20, ,True)
'
'Exit sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim Rs As DAO.Recordset
Dim xlapp As Excel.Application
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim iCols As Integer
Dim ahora As Single
Dim fso As Scripting.FileSystemObject
Dim autoName As Boolean
Dim n As Long
Dim sig As Integer, Contador As Integer
'Las líneas comentadas pertenecen a los controles del test que hemos realizado
'Dim StartTime As Double, EndTime As Double
   
On Error GoTo lbError
    
'Comprueba que no exista el fichero y si existe, lo borra
Set fso = New Scripting.FileSystemObject
    
    If fso.FileExists(strFilename) Then
        fso.DeleteFile strFilename, True
    End If
     
'Abre el recordset que le hemos pasado
Set Rs = CurrentDb.OpenRecordset(strSQL)
'Recupera el total de registros que se van a trasnferir
    If Not (Rs.BOF And Rs.EOF) Then
        DoEvents
        Rs.MoveLast
            sig = Rs.RecordCount
        Rs.MoveFirst
    End If
    
   
    autoName = True
'Crea un nuevo objeto Excel
Set xlapp = New Excel.Application
'Aplica las propiedades según el parámetro boShExcel que le hemos pasado
    With xlapp
        .DisplayStatusBar = boShExcel
        .EnableEvents = boShExcel
        .DisplayAlerts = boShExcel
        .Visible = boShExcel
    End With
'Añade un nuevo libro
Set wb = xlapp.workbooks.Add
'Borra todas las hojas del nuevo libro excel, excepto 1
While wb.sheets.Count > 1
    wb.sheets(wb.sheets.Count).Delete
Wend
'Crea el ojeto hoja
Set ws = wb.sheets(1)

'Cambia el nombre de la hoja
If strSheetName <> "" Then
    ws.Name = Trim(Left(strSheetName, 31))
ElseIf autoName Then
    ws.Name = Trim(Left(fso.GetBaseName(strFilename), 31))
End If

'Crea la primera línea como cabecera con los nombres de los campos
For iCols = 0 To Rs.Fields.Count - 1
    DoEvents
    ws.Cells(1, iCols + 1).Value = Rs.Fields(iCols).Name
    Select Case Rs.Fields(iCols).Type
        Case dbDate
            ws.Columns(iCols + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        Case dbDecimal
            ws.Columns(iCols + 1).NumberFormat = "0"
        Case Else
            ws.Columns(iCols + 1).NumberFormat = "@"
    End Select
Next

DoEvents
'Recorre el recordset y va enviando los paquetes de datos según el rango que hemos marcado según el parámetro pasos
If Not (Rs.BOF And Rs.EOF) Then
    Rs.MoveFirst
'    Debug.Print Time
    For n = 1 To sig Step pasos
'        StartTime = Timer

        DoEvents
        Contador = Contador + 1
        If boShExcel Then xlapp.ScreenUpdating = False
        ws.Range("A" & n + 1).CopyFromRecordset Rs, pasos
        
        If n = 1 Then
            xlapp.ScreenUpdating = True
        End If
            
'EndTime = Timer
'
'Debug.Print "20 registros: " & FormatNumber((EndTime - StartTime), 2, vbFalse, vbFalse, vbFalse) & "s"
'
'Debug.Assert Not Contador = 5

    Next
    ws.Range("A" & sig).Select
End If
    
'Ejecuta algunos arreglos estéticos como "TableStyle"
ws.ListObjects.Add(xlSrcRange, ws.UsedRange, , xlYes).Name = "Tabla1"
ws.ListObjects("Tabla1").TableStyle = "TableStyleMedium2"
 
For iCols = 0 To Rs.Fields.Count - 1
    DoEvents
'Cambia los formatos
    Select Case Rs.Fields(iCols).Type
        Case dbDate
            ws.Columns(iCols + 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End Select
    ws.Cells(1, iCols + 1).Select
'Autoajusta las columnas
    ws.Columns(iCols + 1).EntireColumn.AutoFit
Next
    
If boShExcel Then xlapp.ScreenUpdating = True

ws.Cells(1, 1).Select
    
With xlapp
    .DisplayStatusBar = boShExcel
    .EnableEvents = boShExcel
    .DisplayAlerts = boShExcel
    .ScreenUpdating = boShExcel
End With
    
ws.Cells(1, 1).Select
    
'Graba el fichero
wb.SaveAs FileName:=strFilename
    
GoTo lbFinally
    
lbError:
MsgBox Err & vbCrLf & Error$

Resume
    
'Cierra todos los objetos
lbFinally:
On Error Resume Next
    
    Rs.Close
Set Rs = Nothing
    
Set ws = Nothing
    
    wb.Close
Set wb = Nothing
        
    xlapp.Quit
Set xlapp = Nothing
    
On Error GoTo 0
    
End Sub

