Public Sub CrearAccesoDirecto()
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-crear-un-acceso-directo-a-mi-programa/
'                     Destello formativo 388
'--------------------------------------------------------------------------------------------------------
' Título            : CrearAccesoDirecto
' Autor             : Luis Viadel | luisviadel@access-global.net
' Creado            : desconocido
' Idea original     : https://www.elguille.info/vb/ejemplos/crear_links.htm
' Propósito         : crear un acceso directo a mi programa en tiempo de ejecución
'--------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test.
'                     Copiar el bloque siguiente al portapapeles y pega en el editor de VBA.
'
'Sub CrearAccesoDirecto_Test()
'
'   Call CrearAccesoDirecto
'
'End Sub
'--------------------------------------------------------------------------------------------------------
Dim WScript As Variant
Dim WLink As Variant
Dim escritorio As String, targetLink As String
Dim milink As String, mipath As String, miicono As String

'Localizamos la carpeta escritorio del equipo
    Set WScript = CreateObject("WScript.Shell")
'Podríamos colocarlo en otras carpetas especiales (Programs, StartMenu, StartUp, MyDocuments,...),
'indicando aquí su nombre, e incluso, preguntando al usuario la carpeta dónde quiere colocar el acceso directo.
        escritorio = WScript.SpecialFolders("Desktop")
'Establecemos las propiedades del acceso directo que vamos a crear
        targetLink = "C:\Users\luisv\OneDrive\Escritorio\MyApp\MyApp(00.21).accdb"
        milink = "Mi_App.lnk"
        mipath = escritorio & "\" & milink
        miicono = "C:\Users\luisv\OneDrive\Escritorio\MyApp\Galería\Iconos de la aplicación\Logo_MyApp.ico"
        
'Creamos el acceso directo
        Set WLink = WScript.CreateShortcut(mipath)
            With WLink
'Establecemos algunas propiedades
                .Targetpath = targetLink
                .Description = "Mi App personal"
                .iconlocation = miicono
'               .WorkingDirectory = "Mi directorio de trabajo"
'               .Arguments = "C:/.../MiApp.txt"
'               .Hotkey  = "Ctrl+Alt+..."
'               .WindowStyle   = 1
                .Save
            End With
        Set WLink = Nothing
    Set WScript = Nothing

End Sub