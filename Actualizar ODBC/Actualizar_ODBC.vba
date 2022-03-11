Public Function PruebaConexion() As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/actualizar-una-conexion-odbc-en-tiempo-de-ejecucion
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : PruebaConexion
' Autor original    : Luis Viadel
' Actualizado       : 10/03/2022
' Propósito         : actualizar una conexión
' Retorno           : Verdadero / falso según el resultado obtenido
' Argumento/s       : Si es una única conexión, no sería necesario pasar ningún argumento. Pero se puede pasar la cuenta de donde viene,
'                     por ejemplo, incluso se pueden pasar todos los elementos de la cadena de conexión (dsn, pass, ...)
'                     Parte                      Modo                    Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     NOMBRE_DEL_ARGUMENTO       Obligatorio
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Importante        : la cadena de conexión es para una base de datos de PostgreSQL. Para adaptar la cadena a la base da datos consultar la URL
' Más información   : https://www.connectionstrings.com/
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub pruebaconexion_test()
'Dim funciona as boolean
'
'
'   funciona=PruebaConexion
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim tdfCurrent As DAO.TableDef
Dim tdfLinked As TableDef
Dim strConnectionString As String, NombreDSN As String
Dim tableOld As String, tableNew As String
Dim I As Integer
On Error GoTo LinErr
strConnectionString = "ODBC;DSN=" & Midsn & ";" & _
    "DATABASE=" & Mibd & ";" & _
	"SERVER=" & Miserver & ";" & _
	"PORT=" & Miport & ";" & _
	"UID=" & Miuser & ";" & _
	"PWD=" & Mipass
'Debug.Print strConnectionString
'Conectamos las tablas
For Each tdfCurrent In CurrentDb.TableDefs
	If Len(tdfCurrent.Connect) > 0 Then
		If UCase$(Left$(tdfCurrent.Connect, 5)) = "ODBC;" Then
			If Left(tdfCurrent.NAME, 1) = "~" Then GoTo LinNext
			'Revisamos todas las tablas de la matriz
			For I = LBound(TableName) To UBound(TableName)
				If LCase(tdfCurrent.NAME) = TableName(I) Then GoTo LinGraba
			Next I
			GoTo LinNext
LinGraba:
			tableOld = LCase(tdfCurrent.NAME)
			TableError = tableOld
			'Siel nombre de la tabla contiene el schema, lo incluimos. En el ejemplo es "public"
			tableNew = "public_" & LCase(tableOld)
			'Se puede incluir una pequeña pausa de código si no se ejecuta correctamente
			Set tdfLinked = CurrentDb.CreateTableDef(tableNew)
			tdfLinked.Connect = strConnectionString
			tdfLinked.SourceTableName = tableOld
			CurrentDb.TableDefs.Append tdfLinked
			Set tdfLinked = Nothing
			DoCmd.DeleteObject acTable, tableOld
			DoCmd.Rename tableOld, acTable, tableNew
		End If
	End If
LinNext:
Next
PruebaConexion = True
Exit Function
LinErr:
'Aquí tu tratamiento de errores
PruebaConexion = False
End Function