Public sub GraficoTareas()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/crear-un-grafico-con-google-chart-api/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GraficoTareas
' Autor original    : Luis Viadel
' Actualizado       : julio 2020
' Propósito         : dibujar un gráfico dinámico en un formulario de Access con la API de Google charts
' Referencias       : Microsoft Scripting Runtime (c:\Windows\SysWOW64\scrrun.dll)/>
'                   : Microsoft XML, v6.0 (C:\windows\SysWOW64\msxml6.dll)
' Importante        : Utiliza el objeto WebBrowser que se descataloga en mayo 2022
' Más información   : https://developers.google.com/chart/interactive/docs
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Definimos la variables
Dim objHTML As Object
Dim FSO As Object       'Scripting.FileSystemObject
Dim strHTML As String
Dim ruta As String
Dim rstTable  As DAO.Recordset
'La variable strHTML contendrá la cadena del código que nos proporciona la API para poder dibujar el gráfico.
strHTML = "<html>" & vbCrLf
strHTML = strHTML & "  <head>" & vbCrLf
strHTML = strHTML & "    <script type=""text/javascript"" src=""https://www.google.com/jsapi""></script>" & vbCrLf
strHTML = strHTML & "    <script type=""text/javascript""[>]" & vbCrLf
strHTML = strHTML & "" & vbCrLf
strHTML = strHTML & "      google.load('visualization', '1', {packages:['corechart','line']});" & vbCrLf
strHTML = strHTML & "" & vbCrLf
strHTML = strHTML & "      google.setOnLoadCallback(drawChart);" & vbCrLf
strHTML = strHTML & "" & vbCrLf
strHTML = strHTML & "      function drawChart() {" & vbCrLf
strHTML = strHTML & "" & vbCrLf
'Aquí hay que pasarle los datos de los campos de nuestra tabla o consulta
strHTML = strHTML & "      var data = google.visualization.arrayToDataTable([" & vbCrLf
'Ponemos la cabecera
strHTML = strHTML & "      ['" & rstTable("Fecha").NAME & "', '" & rstTable("Tareas").NAME & "', {role:'tooltip'}],"
'Recorremos los registros para ir incorporándolos al gráfico
Set rstTable = CurrentDb.OpenRecordset("tareas")
Do Until rstTable.EOF
strHTML = strHTML & vbCrLf
strHTML = strHTML & "      ['" & rstTable("FechaTarea") & "'," & rstTable("nTareas") & ", '" & rstTable("FechaTarea") & ": " & Format(rstTable("nTareas"), "General number") & " Tareas'],"
rstTable.MoveNext
Loop
rstTable.Close
Set rstTable = Nothing
'Terminamos la construcción del código según la API
strHTML = strHTML & vbCrLf
strHTML = strHTML & "      ]);" & vbCrLf
strHTML = strHTML & "" & vbCrLf
'Aquí podemos añadir las variables del gráfico (tamaño, efecto, leyensa, etc...)
strHTML = strHTML & "      var options = {" & vbCrLf
strHTML = strHTML & "                       title:'none'," & vbCrLf
strHTML = strHTML & "                       animation:{startup: true, duration: 2000}," & vbCrLf
strHTML = strHTML & "                       legend: 'none'," & vbCrLf
strHTML = strHTML & "                       width: 400," & vbCrLf
strHTML = strHTML & "                       height: 100," & vbCrLf
strHTML = strHTML & "                       chartArea:{left:40,top:10,width:'90%',height:'85%'}," & vbCrLf
strHTML = strHTML & "                       colors:['#5F8AC3']," & vbCrLf
strHTML = strHTML & "                       hAxis:{Title: 'none'}," & vbCrLf
strHTML = strHTML & "                       vAxis:{Title: 'none'}" & vbCrLf
strHTML = strHTML & "                     };" & vbCrLf
strHTML = strHTML & "      var chart = new google.visualization.LineChart(document.getElementById('chart_div'));" & vbCrLf
strHTML = strHTML & "      chart.draw(data, options);" & vbCrLf
strHTML = strHTML & "      };" & vbCrLf
strHTML = strHTML & "      </script>" & vbCrLf
strHTML = strHTML & "  </head>" & vbCrLf
strHTML = strHTML & "    <body>" & vbCrLf
strHTML = strHTML & "      <div id=""chart_div"" style=""width: 800px; height: 200px;""></div>" & vbCrLf
strHTML = strHTML & "    </body>" & vbCrLf
strHTML = strHTML & "</html>" & vbCrLf
'Creamos el archivo, en el directorio de la DB en una carpeta temporal (temp), con el nombre de "tareastemp.html"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objHTML = FSO.CreateTextFile(Application.CurrentProject.Path & "\Temp\tareastemp.html", True)
objHTML.Write strHTML
Set objHTML = Nothing
Set FSO = Nothing
'Vaciamos el control
Form_formEjemplo.ctrlExplorador.ControlSource = ""
'Arreglamos la cadena de la ruta para pasársela al explorador
ruta = Replace(Application.CurrentProject.Path, ":", "$")
ruta = Replace(ruta, "\", "/")
'Cargamos el HTML en el control
Form_formEjemplo.ctrlExplorador.ControlSource = "='file://127.0.0.1/" & ruta & "/Temp/tareastemp.html'"
End function
