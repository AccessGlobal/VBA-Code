Option Compare Database
Option Explicit

Public Sub lstAlign(lstBox As ListBox, lstColumn As Long, ncampos As Long)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-alinear-campo-de-listbox-a-la-derecha/
'					  Destello formativo 341
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : lstAlign
' Autor             : Luis Viadel | luisviadel@access-global.net
' Creado            : 15/06/2023
' Propósito         : alinear a la derecha las columnas numéricas de un listbox
' Más información   : https://access-global.net/aprende-wizhook-con-colin-riddington/
'                     https://access-global.net/vba-crear-y-manipular-una-tabla-en-tiempo-de-ejecucion/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Adáptalo con tus datos y pulsa F5 para ver su funcionamiento.
'
'Public Sub Micampo_AfterUpdate()
'
'    lstAlign Me.MiListBox, nColumna, Me.MiListBox.ColumnCount
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim sizeColumns As String, strSQL As String, strSQL1 As String
Dim sizeCol As Long, valItem As Long
Dim matWidth, matItems
Dim num As Integer, intWhere As Integer
Dim tbl As Object
Dim dbs As DAO.Database
Dim rstTable As DAO.Recordset
Dim ajust As Long
Dim spcBlank As Long
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

'Para poder utilizar de nuevo la tabla temporal, debemos liberar el listbox
    lstBox.RowSource = ""
'Comprobamos si la tabla temporal ya ha sido creada para borrarla
'Este código es necesario para la primera vez
    For Each tbl In CurrentData.AllTables
        If tbl.name = "tempTable" Then
            DoCmd.DeleteObject acTable, "tempTable"
        End If
    Next tbl
'Traemos las dimensiones de las columnas
    sizeColumns = lstBox.ColumnWidths
'Grabamos en una matriz las dimensiones una a una
    matWidth = Split(sizeColumns, ";")
'Recorremos la matriz buscando el tamaño de la columna solicitada
    For num = 0 To UBound(matWidth)
        If num = lstColumn Then
            sizeCol = matWidth(lstColumn)
        End If
    Next
'ajust es una variable de ajuste porque los cálculos no son exactos.
'Se obtiene realizando pruebas, no tiene ningún fundamento más que ajustar los resultados
    If sizeCol > 500 And sizeCol < 1000 Then
        ajust = 0
    End If
    If sizeCol > 1000 And sizeCol < 2000 Then
        ajust = 4
    End If
    If sizeCol > 2000 And sizeCol > 3000 Then
        ajust = 6
    End If
    If sizeCol > 3000 Then
        ajust = 7
    End If

'Creamos una tabla temporal donde guardamos los datos del listbox convertidos en texto (véase Destello 191)
    For num = 1 To lstBox.ColumnCount
        strSQL = strSQL & " campo" & num & " CHAR,"
    Next
    
    strSQL = "CREATE TABLE tempTable (" & Left(strSQL, Len(strSQL) - 1) & ");"
    Set dbs = CurrentDb
        dbs.Execute strSQL
'Rellena la nueva tabla con los datos del listbox
        strSQL = "SELECT idprodut, produtcod, produtnom, precio, activo FROM productos;"

        For num = 1 To ncampos
            If num = 1 Then
                strSQL1 = "campo" & num & ", "
            ElseIf num = ncampos Then
                strSQL1 = strSQL1 & "campo" & num
            Else
                strSQL1 = strSQL1 & "campo" & num & ", "
            End If
        Next num
        
        dbs.Execute " INSERT INTO tempTable (" & strSQL1 & ") " & strSQL
    Set dbs = Nothing
    
'Otras posibilidades que mejorarían el ejemplo
'1. Recorremos la tabla para comprobar los decimales en los números (Se podrían unificar los números)
'2. A los campos moneda les podemos incorporar el símbolo de moneda
'...
'Ordenamos el campo a la derecha o la izquierda según la decisión del usuario
'Calculamos cuantos Twips tiene un carácter en blanco
    On Error Resume Next
    
    WizHook.key = 51488399
'Para ordenar, rellenaremos de espacios en blanco el campo
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

    strSQL = "Campo" & lstColumn + 1
    Set rstTable = CurrentDb.OpenRecordset("SELECT " & strSQL & " FROM tempTable")
        Do Until rstTable.EOF
            rstTable.Edit
                spcBlank = (sizeCol / wzdx) - (2 * (Len(c(rstTable.Fields(strSQL))))) - ajust
                rstTable.Fields(strSQL).value = RTrim(Space(spcBlank) & Trim(rstTable.Fields(strSQL)))
            rstTable.Update
            rstTable.MoveNext
        Loop
    Set dbs = Nothing
    
    lstBox.RowSource = "SELECT * FROM temptable"

End Sub
