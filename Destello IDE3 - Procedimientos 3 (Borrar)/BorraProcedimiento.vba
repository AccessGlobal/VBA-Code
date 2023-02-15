'Añadir en el evento "Al hacer click" de un botón
Private Sub btnBorrar_Click()
Dim vbc As VBIDE.VBComponent
Dim strProc As String, strMod As String
Dim proctipo As String, intStr As Integer
Dim lngProcTyp As Long
Dim lngStartLine As Long
Dim i As Integer

'Seleccionamos el procedimiento que el usuario ha seleccionado en el listbox
    strProc = Me.lstProc.Value
    
'Necesitaríamos un control para comprobar si se trata de un módulo o un procedimiento
    intStr = InStr(1, strProc, "(")
'Extraemos el nombre
    strProc = Trim(Right(Left(strProc, intStr - 1), Len(Left(strProc, intStr - 1)) - 9))
    
'Necesitamos el nombre del módulo como segundo argumento
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
                
        With vbc.CodeModule
            strMod = .CodePane.CodeModule
            
            lngStartLine = .CountOfDeclarationLines + 1
          
            For i = lngStartLine To .CountOfLines
'Obtenemos el tipo de procedimiento con ProcOfLine a través del número de línea
                 If strProc = .ProcOfLine(lngStartLine, lngProcTyp) Then
                    Call BorraProcedimiento(strMod, strProc, lngProcTyp)
                    MsgBox "El procedimiento ha sido borrado con éxito"
                    Exit Sub
                End If
                lngStartLine = lngStartLine + 1
            Next i
           
        End With
        
    Next vbc
    
    MsgBox "No he encontrado el procedimiento"
        
End Sub

'Código a incorporar en un módulo estándar
Public Sub BorraProcedimiento(strModuelName As String, strProzedur As String, strProcTyp As Long)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-borrar-procedimiento/
'                     Destello formativo 266
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : BorraProcedimiento
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : borra un procedimiento concreto
' Argumentos        : La sintaxis del procedimiento consta de los siguientes argumentos:
'                     Parte           Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strModuelName   Obligatorio      Nombre del módulo que contiene el procedimiento
'                     strProzedur     Obligatorio      Nombre del procedimiento que queremos borrar
'                     strProcTyp     Obligatorio       Tipo del procedimiento que queremos borrar
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente en el
'                     evento de un botón. No olvides dar valor a las variables strModuloName (Nombre del módulo donde vas a crear el
'                     procedimiento) y strVBACode (que contiene el procedimiento)
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub MiBoton_Click()
'
'    Call BorraProcedimiento("Nombre del módulo", "Nombre del procedimiento", "Tipo de porcedimiento")
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim lngFirstLine As Long
Dim lngLastLine As Long

    With Application.VBE.ActiveVBProject.VBComponents(strModuelName).CodeModule
           
        lngFirstLine = .ProcStartLine(strProzedur, strProcTyp)
        lngLastLine = .ProcCountLines(strProzedur, strProcTyp)
        .DeleteLines lngFirstLine, lngLastLine
      
    End With

End Sub
