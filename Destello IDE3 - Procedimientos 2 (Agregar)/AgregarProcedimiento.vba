'Añadir en el evento "Al hacer click" de un botón
Private Sub btnAdd_Click()
Dim strVBACode As String
Dim strProc As String
Dim intStr As Integer

'Seleccionamos el módulo que el usuario ha seleccionado en el listbox
    strProc = Me.lstProc.Value
    
    intStr = InStr(1, strProc, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strProc = strProc
    Else
        strProc = Left(strProc, intStr - 1)
    End If
    
    If strProc = "" Then
        Exit Sub
    End If

    strVBACode = "Public Sub Procedimiento_Prueba ()" & _
                vbNewLine & _
                vbNewLine & _
                "    msgbox" & """Es una prueba de código remoto""" & _
                vbNewLine & _
                vbNewLine & _
                "End Sub"
                
    Call AgregaProcedimiento(strProc, strVBACode)
    
    MsgBox "Ya he añadido el procedimiento"

End Sub

'Añadir en un módulo estándar
Option Compare Database
Option Explicit

Public Sub AgregaProcedimiento(strModuelName As String, strVBACode As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-agregar-procedimientos/
'                     Destello formativo 264
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AgregaProcedimiento
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : agregar un nuevo porcedimiento a nuestro programa
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa443959(v=vs.60)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente en el
'                     evento de un botón. No olvides dar valor a las variables strModuloName (Nombre del módulo donde vas a crear el
'                     procedimiento) y strVBACode (que contiene el procedimiento)
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub MiBoton_Click()
'Dim strVBACode As String
'
'    strVBACode = "Public Sub Procedimiento_Prueba ()" & _
'                vbNewLine & _
'                "msgbox" & """Es una prueba de código remoto""" & _
'                vbNewLine & _
'                "End Sub"
'
'    Call AgregaProcedimiento("ModPruebas", strVBACode)
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim lngLine As Long

    With Application.VBE.ActiveVBProject.VBComponents(strModuelName).CodeModule
    
        lngLine = .CountOfLines + 1
        .InsertLines lngLine, strVBACode
        
    End With

End Sub
