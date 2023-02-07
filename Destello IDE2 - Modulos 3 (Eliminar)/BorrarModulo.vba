'Pon este código en el evento "Al hacer click" de un botón
Private Sub btnBorrar_Click()
Dim vbc As VBIDE.VBComponent
Dim strMod As String
Dim intStr As Integer
Dim rest As Integer

    strMod = Me.lstMods.Value
    
    intStr = InStr(1, strMod, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strMod = strMod
    Else
        strMod = Left(strMod, intStr - 1)
    End If
    
    If strMod = "" Then
        Exit Sub
    End If
'Pregunta de seguridad para evitar error de borrado
    Set vbc = Application.VBE.ActiveVBProject.VBComponents(strMod)
        rest = MsgBox("¿Quieres eliminar el módulo '" & vbc.Name & "'?", vbCritical + vbOKCancel)
        Do
            If rest = 1 Then
'                Application.VBE.ActiveVBProject.VBComponents.Remove vbc
                BorrarModulo vbc
                ListadoModulos Me
                GoTo LinSalir
            ElseIf rest = 2 Then
                GoTo LinSalir
            End If
        Loop
LinSalir:
   Set vbc = Nothing
    
End Sub


'Pon este código en un módulo estándar
Option Compare Database
Option Explicit

Public Sub BorrarModulo(strModuleName As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-borrar-modulo/
'                     Destello formativo 260
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : BorrarModulo
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : elimina un módulo de nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/remove-method-vba-add-in-object-model
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    BorrarModulo strmoduleName
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim vbc As VBIDE.VBComponent

    Set vbc = Application.VBE.ActiveVBProject.VBComponents(strModuleName)
            
        Application.VBE.ActiveVBProject.VBComponents.Remove vbc
                
    Set vbc = Nothing
           
End Sub
