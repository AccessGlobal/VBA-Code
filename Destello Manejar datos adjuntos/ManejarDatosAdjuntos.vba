'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-como-manejar-el-tipo-datos-adjuntos-mediante-vba/
'                     Destello formativo 382
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Autor             : Luis Viadel | luisviadel@access-global.net
' Fecha             : 11/2023
' Propósito         : Manejar mediante DAO el tipo de datos de Access "Datos adjuntos" (Se incluye BBDD de ejemplo)
'                     1 Grabar un fichero en un campo tipo "Datos adjuntos"
'                       Private Sub btnGrabar_Click()
'                     2 Extraer el fichero contenido en un campo tipo "datos adjuntos"
'                       Private Sub btnRecuperar_Click()
'                     3 Eliminar el contenido de un campo tipo "datos adjuntos" sin borrar todo el registro
'                       Private Sub btnEliminar_Click()
'
'						
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btnGrabar_Click()
Dim fichero As String
Dim Path_inicial As String
Dim rstTable As DAO.Recordset
Dim rstData As DAO.Recordset
   
    Path_inicial = "C:\Cow Technologies\Access global\Destellos formativos\Destello 382\"

    fichero = mcFileDialog(Path_inicial)
    
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM datos WHERE iddatos=1")
        rstTable.Edit
        Set rstData = rstTable.Fields("datosAdj").Value
            rstData.AddNew
                rstData.Fields("FileData").LoadFromFile fichero
            rstData.Update
        Set rstData = Nothing
        rstTable.Update
    Set rstTable = Nothing
    
End Sub

Private Sub btnEliminar_Click()
Dim rstTable As DAO.Recordset
Dim rstData As DAO.Recordset
      
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM datos WHERE iddatos=1")
        rstTable.Edit
        Set rstData = rstTable.Fields("datosAdj").Value
            rstData.Delete
        Set rstData = Nothing
        rstTable.Update
    Set rstTable = Nothing

End Sub

Private Sub btnRecuperar_Click()
Dim rstTable As DAO.Recordset
Dim rstData As DAO.Recordset
Dim ruta As String

    ruta = Application.CurrentProject.Path & "\ImgTemp.jpg"
    
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM datos WHERE iddatos=1")
        Set rstData = rstTable.Fields("datosAdj").Value
            While Not rstData.EOF
                rstData.Fields("FileData").SaveToFile ruta
                GoTo Exitsub
            Wend
        Set rstData = Nothing
    Set rstTable = Nothing

Exitsub:
    Call ShellExecute(0&, "open", ruta, 0&, vbNullString, 1&)
    
End Sub
