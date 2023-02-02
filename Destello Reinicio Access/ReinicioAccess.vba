Option Compare Database
Option Explicit

' Tiempo de espera establecido en 60 iteraciones, después de lo cual el archivo por lotes debe eliminarse a sí mismo
Private Const TIMEOUT = 60

Public Sub Restart(compact As Boolean)
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-reiniciar-ms-access
' Fuente original   : http://blog.nkadesign.com/microsoft-access/
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Restart
' Autor original    : Renaud Bompuis
'                     Licensed under the Creative Commons Attribution License
'                     http://creativecommons.org/licenses/by/3.0/
'                     http://creativecommons.org/licenses/by/3.0/legalcode
' Adaptado          : Luis Viadel
' Creado            : 2008-2009
' Propósito         : reiniciar esta aplicación
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte        Modo            Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------------
'                     frm          Obligatorio     Formulario que hace la llamada
'                     filtro       Opcional        tipo del módulo que queremos listar
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Cómo funciona     : Primero, creamos un pequeño archivo por lotes al que le pasamos la ruta y la extensión
'                     del ejecutable de Access, la base de datos actual y su extensión de archivo de bloqueo.
'                     Luego ejecutamos ese script y cerramos la aplicación.
'                     El script supervisa la presencia del archivo de bloqueo.
'                     Tan pronto como Access elimine el archivo de bloqueo:
'                      - si se dio la opción Compactar, primero compactamos la base de datos
'                      - abrimos la base de datos de nuevo.
'                     El script se eliminará automáticamente
' Referencias       : http://www.dx21.com/HOME/ARTICLES/P2P/ARTICLE.ASP?CID=12
'                     http://malektips.com/dos0017.html
'                     http://www.catch22.net/tuts/selfdel.asp
'-------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pégalo en el editor de VBA. Descomenta la línea que interese y pulsa F5 para ver su funcionamiento.
'
'Sub Restart()
'
'   restart (True)
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim scriptpath As String
Dim S As String
Dim dbname As String, ext As String, lockext As String, accesspath As String
Dim idx As Integer
Dim intFile As Integer
    
' Full path del fichero temporal
    scriptpath = Application.CurrentProject.FullName & ".dbrestart.bat"
    
' Si el script ya existe, verifica que no sea un remanente antiguo
' que ha pasado su tiempo de espera.
    If Dir(scriptpath, vbNormal) <> "" Then
        If DateAdd("s", TIMEOUT * 2, FileDateTime(scriptpath)) < Date Then
' Pasamos el doble del tiempo de espera del archivo por lotes, dando suficiente tiempo para que ' se ejecute, por lo que si todavía está allí, lo más probable es que sea un tipo
            Kill scriptpath
        Else
'El tiempo de espera no ha expirado más allá del límite aceptable, por lo que probablemente sea
' aún activo, solo intente cerrar la aplicación nuevamente
            Application.Quit acQuitSaveAll
            Exit Sub
    End If
End If
    
' Construimos el archivo por lotes
' Ten en cuenta que el valor TIMEOUT solo se usa como un contador de bucle y
' realmente no contamos el tiempo transcurrido en el archivo por lotes.
' El comando ping tarda un tiempo en cargarse e iniciarse y aunque
' establecemos su tiempo de espera en 100 ms, tardará mucho más que eso en
' ejecutar.
' Si nos han pedido que compactemos la base de datos, lanzamos la base de datos
' usando el modificador de línea de comando /compacto
    S = S & "SETLOCAL ENABLEDELAYEDEXPANSION" & vbCrLf
    S = S & "SET /a counter=0" & vbCrLf
    S = S & ":CHECKLOCKFILE" & vbCrLf
    S = S & "ping 0.0.0.255 -n 1 -w 100 > nul" & vbCrLf
    S = S & "SET /a counter+=1" & vbCrLf
    S = S & "IF ""!counter!""==""" & TIMEOUT & """ GOTO CLEANUP" & vbCrLf
    S = S & "IF EXIST ""%~f2.%4"" GOTO CHECKLOCKFILE" & vbCrLf
    
    If compact Then
    
        S = S & """%~f1"" ""%~f2.%3"" /compact" & vbCrLf
    
    End If
    
    S = S & "start "" "" ""%~f2.%3""" & vbCrLf
    S = S & ":CLEANUP" & vbCrLf
    S = S & "del %0"
    
' Escribe el fichero por lotes
    intFile = FreeFile()
    
    Open scriptpath For Output As #intFile
    
    Print #intFile, S
    
    Close #intFile
            
' Crea los argumentos a pasar al script
' Aquí le pasamos la ruta completa a la base de datos menos la extensión que pasamos por separado
' esto se hace para que podamos reconstruir fácilmente la ruta al archivo de bloqueo en el script.
' La extensión del archivo de bloqueo también se pasa como un tercer argumento.

' Obtenga la ruta al ejecutable msaccess, donde sea que esté
    accesspath = SysCmd(acSysCmdAccessDir) & "msaccess.exe"
    
' Encuentra la extensión, comenzando desde el final
    For idx = Len(CurrentProject.FullName) To 1 Step -1
        If Mid(CurrentProject.FullName, idx, 1) = "." Then Exit For
    Next idx

    dbname = Left(CurrentProject.FullName, idx - 1)
    ext = Mid(CurrentProject.FullName, idx + 1)
    
    lockext = "laccdb"
    
' Llama al fichero por lotes
    S = """" & scriptpath & """ """ & accesspath & """ """ & dbname & """ " & ext & " " & lockext
    Shell S, vbHide
    
' Cierra la aplicación
Application.Quit acQuitSaveAll

End Sub
