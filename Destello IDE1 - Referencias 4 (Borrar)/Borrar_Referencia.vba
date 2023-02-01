Option Compare Database
Option Explicit

Public Sub BorrarReferencia(strReference As String)
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-eliminar-referencias
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : BorrarReferencia
' Autor original    : Alba Salvá
' Adaptado          : Luis Viadel
' Creado            : desconocido
' Propósito         : eliminar una referencia del proyecto usando el GUID para la biblioteca de referencia
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------------
'                     strReference    Obligatorio      GUID de la referencia que queremos habilitar
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/office/vba/api/access.references.remove
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub BorrarReferencia_test()
'Dim res as boolean
'Dim strGUID As String
'
'Introduce el GUID de la librería
'   strGUID = "{F5078F18-C551-11D3-89B9-0000F81FE221}" 'Microsoft XML, v6.0
'   res = BorrarReferencia(strGUID)
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim objRef As VBIDE.reference

    On Error GoTo LinError
    
    Set objRef = Application.VBE.ActiveVBProject.References(strReference)
    
        Application.VBE.ActiveVBProject.References.Remove objRef
       
        GoTo LinFinalizar
    
LinError:
        MsgBox "No se puede eliminar la referencia" & vbCrLf & "'" & objRef.name & "'", vbCritical, "Borrar referencia"

LinFinalizar:
    Set objRef = Nothing

End Sub
