'Colocar en un formulario que contiene un cuadro de lista (lstProductos)
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)
Dim strSQL As String

    strSQL = "SELECT productos.idprodut, productos.produtcod, productos.produtnom, productos.produfec, productos.activo FROM productos;"
    
    Me.LstProductos.ColumnCount = 5
    Me.LstProductos.RowSource = strSQL
    Me.LstProductos.ColumnWidths = "0cm;2cm;7cm;2cm;1cm"
    
End Sub

Private Sub LstProductos_DblClick(Cancel As Integer)

    MsgBox "El id del produto que has seleccionado es el " & Me.LstProductos.value
    
End Sub

Private Sub ver02_AfterUpdate()

    cambiarLista
    
End Sub

Private Sub ver03_AfterUpdate()

    cambiarLista
    
End Sub

Private Sub ver04_AfterUpdate()

    cambiarLista
    
End Sub

Private Sub ver05_AfterUpdate()

    cambiarLista
    
End Sub

Private Sub cambiarLista()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-cuadro-de-lista-personalizado/
'                     Destello formativo 340
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : cambiarLista
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 2020
' Adaptado ejemplo  : 2023
' Propósito         : dar al usuario la opción de mostrar u ocultar campos de un cuadro de lista
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el evento "Al cambiar" de una casilla de verificación.
'
'Private Sub CampoDeLista_AfterUpdate()
'
'       cambiarLista
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim anchoColumna As String
Dim nColumnas As Long
Dim strSQL As String

    strSQL = "SELECT productos.idprodut"
    nColumnas = 1
    anchoColumna = "0cm;"
    
     If ver02 = True Then
        nColumnas = nColumnas + 1
        anchoColumna = anchoColumna & "2cm;"
        strSQL = strSQL & ", productos.produtcod"
    End If
   
    If ver03 = True Then
        nColumnas = nColumnas + 1
        anchoColumna = anchoColumna & "7cm;"
        strSQL = strSQL & ", productos.produtnom"
    End If
   
    If Ver04 = True Then
        nColumnas = nColumnas + 1
        anchoColumna = anchoColumna + "2cm;"
        strSQL = strSQL & ", productos.produfec"
    End If
   
    If ver05 = True Then
        nColumnas = nColumnas + 1
        anchoColumna = anchoColumna + "1cm;"
        strSQL = strSQL & ", productos.activo"
    End If
      
    strSQL = strSQL & " FROM productos"
            
    Me.LstProductos.ColumnCount = nColumnas
    Me.LstProductos.RowSource = strSQL
    Me.LstProductos.ColumnWidths = anchoColumna
    
End Sub
