Option Compare Database
Option Explicit


Public Sub AgregaModulo(strModuleName As String, strModuleType As CompType)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-agregar-modulo/
'                     Destello formativo 260
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AgregaModulo
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : agrega un nuevo módulo a nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-vba-add-in-object-model
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ListadoModulosListadoModulos Me,NombreFiltro
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent

    On Error GoTo lbError

    Set vbc = Application.VBE.ActiveVBProject.VBComponents.Add(strModuleType)
        
        If Not vbc Is Nothing Then
            vbc.Name = strModuleName
            Debug.Print vbc.Saved
            
        End If
    
        GoTo lbFinally

lbError:
    If Err = 32813 Then
        MsgBox "El nombre " & strModuleName & " ya está en uso." & vbCrLf & "Intoduzca otro nombre.", vbCritical
    End If
    
    Exit Sub
    
lbFinally:
        On Error GoTo 0
    
    Set vbc = Nothing
    
    MsgBox "Ahora puedes utilizar la función restart del destello 258 - Reiniciar MsAccess para 'Compactar y reparar' y salir y volver a entrar para guardar los cambios", vbInformation
    
'    Restart (True)
    
    
End Sub