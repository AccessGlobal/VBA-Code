'Pon este código en el evento "Al hacer click" de un botón
Private Sub btnRename_Click()
Dim respuesta As String
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

    respuesta = ControlarRespuestaInputBox("Indica el nuevo nombre para el módulo '" & strMod & "'")
    
    If respuesta = "" Then
        MsgBox "Se ha producido un error y no se puede renombrar el módulo", vbExclamation
        Exit Sub
    Else
        RenombraModulo strMod, respuesta
    End If

End Sub

'Pon este código en un módulo estándar
Option Compare Database
Option Explicit

Public Sub RenombraModulo(strModuleName As String, strNewName As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-renombrar-modulo/
'                     Destello formativo 261
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : RenombraModulo
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : cambia el nombre de un módulo de nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte               Modo                   Descripción
'                     strModuleName    Obligatorio        Nombre del módulo que queremos renombrar
'                     strNewName       Obligatorio        Nuevo nombre
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-vba-add-in-object-model
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    RenombraModulo strModuleName, strModuleNewName
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
    
    Application.VBE.ActiveVBProject.VBComponents(strModuleName).Name = strNewName
 
End Sub