Sub TestMacros(strForm As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-comprobar-macros-incrustadas/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : TestMacros
' Autor original    : Alba Salvá
' Creado            : Desconocido
' Propósito         :
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte         Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strForm       Obligatorio    Nombre del formulario que se tiene que analizar
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, rellena los datos de la url y los del
'                     fichero que deseas descargar y pulsa F5 para ver su funcionamiento.
'
'Sub TestMacros_test()
'Dim frm As Object
    
'    For Each frm In CurrentProject.AllForms
'        TestMacros frm.Name
'    Next
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim frm As Form
Dim ctl As Control
Dim prp As Property
    
    
    If FormularioAbierto(strForm) Then DoCmd.Close acForm, strForm
    
    DoCmd.OpenForm strForm, acDesign, , , , acHidden
    Set frm = Forms(strForm)
    For Each ctl In frm.Controls
        For Each prp In ctl.Properties
            On Error Resume Next
            If InStr(prp.Name, "EmMacro") And Left(Trim(prp.Name), 2) = "On" Then
                If Len(Trim(prp.Value & "")) > 0 Then
                    
                    Debug.Print "Formulario: "; strForm
                    Debug.Print "Control: "; ctl.Name
                    Debug.Print "Evento : "; Replace(prp.Name, "EmMacro", "")
                    Debug.Print "Contenido: "; vbCrLf; _
                                "------------------------------------------------------------------------------------------"; vbCrLf; _
                                prp.Value
                    Debug.Print "------------------------------------------------------------------------------------------"
                    Debug.Print
                    Debug.Print
                End If
            End If
        Next
    Next
    DoCmd.Close acForm, strForm, acSaveNo
    
End Sub
