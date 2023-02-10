'Código para el evento "OnClick" de un botón
Private Sub btnExport_Click()
Dim strFile As String, strMod As String
Dim intStr As Integer, rest As Integer
Dim strTipo As String
Dim tipo As Integer
Dim vbc As VBIDE.VBComponent

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
    
'Buscamos el tipo de módulo para añadir el tipo de archivo
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        If vbc.Name = strMod Then
            tipo = vbc.Type
            Exit For
        End If
    
    Next
    
    Select Case tipo
        Case 1 'Módulo estándar"
            strTipo = ".bas"
        Case 2
            strTipo = ".cls"
        Case Else
            strTipo = txt
    End Select
'Indicamos el path completo del fichero al que queremos exportar
    strFile = Application.CurrentProject.Path & "/" & strMod & strTipo
    
    ExportarModulo strMod, strFile
    
End Sub

'Código para incluir en un módulo estándar

Option Compare Database
Option Explicit

Public Sub ExportarModulo(strModuleName As String, strFile As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-modulos-exportar-modulo/
'                     Destello formativo 263
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExportarModulo
' Autor original    : Alba Salvá
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Creado            : febrero 2023
' Propósito         : exporta el módulo a un fichero de texto
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte               Modo                   Descripción
'                     strModuleName    Obligatorio        Nombre del módulo que queremos renombrar
'                     strFile          Obligatorio        Ruta completa del fichero en el que se exportará el módulo.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Funcionamiento    : La función crea un fichero con el nombre y en la dirección que le pasamos como argumento. Si existe un fichero con ese
'                     nombre, lo elimina.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la carga de un formulario.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'
'    ExportarModulo strModuleName, strFile
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
    
    Application.VBE.ActiveVBProject.VBComponents(strModuleName).Export (strFile)
 
End Sub

