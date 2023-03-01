Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                                      ByVal lpOperation As String, _
                                                                                      ByVal lpFile As String, _
                                                                                      ByVal lpParameters As String, _
                                                                                      ByVal lpDirectory As String, _
                                                                                      ByVal nShowCmd As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-shellexecute
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ShellExecute
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : marzo 2014
' Propósito         : ejecutar cualquier archivo desde la API
' Retorno           : sin retorno, salvo que deseemos una confirmación
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte         Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     hwnd          Obligatorio      Identifica la ventana principal
'                     lpOperation   Obligatorio      Acción que se desea realizar al ejecutar la función
'                                                    Parámetros de lpOperation:
'                                                    Open       : abre cualquier documento, incluso una carpeta o directorio
'                                                    Print      : imprime un archivo, si es imprimible, si no lo abre y funciona como ""open"
'                                                    Explore    : se utiliza para carpetas y directorios
'                                                    Play       : reproduce un archivo de sonido
'                                                    Properties : muestra las propiedades del archivo
'                                                    0& o NULL  : toma la acción por defecto (open)
'                     lpFile        Obligatorio      Fichero que queremos ejecutar
'                     lpParameters  Obligatorio      En el caso de fichero ejecutables indicamos los parámetros, en ficheros de texto, 0& o vbNullString
'                     lpDirectory   Obligatorio      Directorio de trabajo por defecto
'                     nShowCmd      Obligatorio      Cómo se verá la aplicación cuando la ejecutemos
'                                                    Parámetros de nShowCmd:
'                                                    0 Oculta la ventana que se está activando y pone el foco en otra ventana.
'                                                    1 Muestra la ventana, restaurándola si se encuentra maximizada o minimizada
'                                                    2 Muestra la ventana en forma de icono .
'                                                    3 Activa la venta maximizada
'                                                    4 Muestra una ventana en su tamaño más reciente y posición. La ventana actual continua activa
'                                                    5 Muestra la ventana en la misma posición y tamaño que tiene actualmente.
'                                                    6 Minimiza la ventana
'                                                    7 La ventana se muestra como un icono. La ventana que está activa, permanece activa
'                                                    8 La ventana se muestra en su estado actual. La ventana que está activa, permanece activa
'                                                    9 Funciona igual que uno.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub ProbarShellExecute()
'Dim sFile As String
'
'sFile = "C:\Cow Technologies\Desarrollo\Cow Harmony desarrollo\Sound\bell.mp3"
'
'ShellExecute 0&, "Open", sFile, 0&, vbNullString, 1&
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------

