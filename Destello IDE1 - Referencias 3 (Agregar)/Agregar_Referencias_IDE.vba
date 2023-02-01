Option Compare Database
Option Explicit

Public Function AddReferenceFromFile(strAdd As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-agregar-referencias
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AddReferenceFromFile
' Autor original    : Alba Salvá
' Adaptado          : Luis Viadel
' Creado            : desconocido
' Propósito         : agregar una referencia al proyecto usando ruta y nombre del fichero para la biblioteca de referencia
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------------
'                     strAdd         Obligatorio       Dirección de la referencia que queremos agregar
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/office/vba/api/access.references.addfromfile
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub AddReferenceFromFile_test()
'Dim retval As Boolean
'Dim strAdd As String
'
'Introduce el path de la librería
'    strAdd = "C:\Windows\SysWOW64\msxml6.dll" 'Microsoft XML, v6.0
'
'    retval = AddReferenceFromFile(strAdd)

'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------

Dim objRefs As VBIDE.References ' Colección

    Set objRefs = Application.VBE.ActiveVBProject.References

        On Error GoTo LinError
    
' Sintaxis del método   : AddFromGuid(Guid, Mayor, Menor)
' Argumentos del método:
'   GUID
'   Major
'   Minor
' En el caso de Microsoft XML, v6.0 (Mayor = 6, Menor = 0)
        objRefs.AddFromFile strAdd
        
        AddReferenceFromFile = True
       
    Set objRefs = Nothing
 
    Exit Function
    
LinError:

    AddReferenceFromFile = False
 
End Function

Public Function AddReferenceFromGUID(strGUID As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-agregar-referencias
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AddReferenceFromGUID
' Autor original    : Alba Salvá
' Adaptado          : Luis Viadel
' Creado            : desconocido
' Propósito         : agregar una referencia al proyecto usando el GUID para la biblioteca de referencia
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'---------------------------------------------------------------------------------------------------------------------------------------------------
'                     strGUID         Obligatorio      GUID de la referencia que queremos habilitar
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/office/vba/api/access.references.addfromguid
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub AddReferenceFromGUID_test()
'Dim res as boolean
'Dim strGUID As String
'
'Introduce el GUID de la librería
'   strGUID = "{F5078F18-C551-11D3-89B9-0000F81FE221}" 'Microsoft XML, v6.0
'   res = AddReference(strGUID)
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------

Dim objRefs As VBIDE.References ' Colección

'Comprobamos que la referencia no está en el poryecto
    If ComprobarReferencias(, strGUID) = True Then
        MsgBox "La referencia ya está en el proyecto"
        AddReferenceFromGUID = False
        Exit Function
    End If

    Set objRefs = Application.VBE.ActiveVBProject.References

        On Error Resume Next
    
' Sintaxis del método   : AddFromGuid(Guid, Mayor, Menor)
' Argumentos del método :
'   GUID
'   Mayor
'   Menor
' En el caso de Microsoft XML, v6.0 : Mayor = 6, Menor = 0
        objRefs.AddFromGuid strGUID, 6, 0
    
' Si quieres elminar la referencia
'        oRefs.Remove (reference)
        AddReferenceFromGUID = True
        
    Set objRefs = Nothing
 
 
End Function
