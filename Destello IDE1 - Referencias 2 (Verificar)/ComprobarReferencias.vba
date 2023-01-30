Public Function ComprobarReferencias(Optional strAdd As String, Optional strGUID As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-referencias-verificar-la-existencia-de-una-referencia
'                     Destello formativo 255
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ComprobarReferencias
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : comprueba si la referencia pasada como argumento ya está referenciada en la aplicación a través de dos de sus propiedades,
'                     la dirección física y el GUID
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strAdd         Opcional      Path completo de la referencia que queremos habilitar
'                     strGUID        Opcional      GUID de la referencia que queremos comprobar
' Nota              : ambos argumentos están marcados como opcionales pero, evidentemente, es necesario por lo menos uno para poder localizarla
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
' Información       : https://learn.microsoft.com/es-es/office/vba/api/access.application.vbe
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : para adaptar este código en tu aplicación simplemente cópialo y pégalo en un módulo estándar
'
'Sub comprobarReferencias_test()
'Dim res As Boolean
'Dim strAdd As String
'Dim strGUID As String
'
'    strAdd="C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
''    strGUID = "{0002E157-0000-0000-C000-000000000046}"
'
'    res = ComprobarReferencias(strAdd)
''    res = ComprobarReferencias(, strGUID)
''    res = ComprobarReferencias(strAdd, strGUID)
'
'    If res Then
'        MsgBox "La referencia existe"
'    Else
'        MsgBox "La referencia no existe"
'    End If
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objRef As VBIDE.reference   ' Referencia
Dim objRefs As VBIDE.References ' Colección

    ComprobarReferencias = False
    
    Set objRefs = Application.VBE.ActiveVBProject.References
     
        For Each objRef In objRefs
        
            If strAdd <> "" Then
                    If objRef.FullPath = strAdd Then
                        ComprobarReferencias = True
                    Else
                        ComprobarReferencias = False
                    End If
            ElseIf strGUID <> "" Then
                    If objRef.Guid = strGUID Then
                        ComprobarReferencias = True
                    Else
                        ComprobarReferencias = False
                    End If
            Else
                MsgBox "Debe indicar al menos un parámetro"
                ComprobarReferencias = False
            End If
            
        Next objRef
     
    Set objRefs = Nothing

End Function
