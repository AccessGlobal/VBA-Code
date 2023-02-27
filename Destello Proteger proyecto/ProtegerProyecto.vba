'Copia y pega este código en un módulo estándar
Public Function ProyectoProtegido() As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-proteger-proyecto/
'                     Destello formativo 274
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ProyectoProtegido
' Autor original    : Alba Salvá
' Creado            : febrero 2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : comprueba si nuestro proyecto está protegido con contraseña
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------

  ProyectoProtegido = CBool(Application.VBE.ActiveVBProject.Protection)

End Function

Public Function PonerPasswordVBA(strDBName As String, strConnect As String, _
                strVBAPWD As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-proteger-proyecto/
'                     Destello formativo 274
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : PonerPasswordVBA
' Autor original    : Alba Salvá
' Creado            : febrero 2023
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Propósito         : establecer la contraseña para el proyecto
' Argumento/s       : La sintaxis de la función consta de los siguientes argumentos:
'                     Parte               Modo                   Descripción
'                     ---------------------------------------------------------------------------------------------------------------------------
'                     strDBName           Obligatorio        Ruta y nombre de la base de datos
'                     strConnect          Obligatorio        No he encontrado información, se puede poner una cadena vacía ("")
'                     strVBAPWD           Obligatorio        Contraseña
'------------------------------------------------------------------------------------------------------------------------------------------------
' Información       : Esta rutina se basa en el uso de la librería indocumentada WizHook.
'                     Aunque hay bastante información que se ha ido obteniendo, aún existen lagunas
'                     de todas las funciones que contiene y/o de los argumentos que se usan
'------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copia el bloque siguiente al
'                     portapapeles y pega en el editor de VBA en la acción que desees.
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Sub PonerPasswordVBA_test()
'Dim respuesta As Variant
'Dim strRespuesta As String
'Dim NewPass As Boolean
'
'    If ProyectoProtegido Then
'        MsgBox "El proyecto ya está protegido"
'    Else
'        respuesta = InputBox("Indica la nueva contraseña para el proyecto")
'
'        If StrPtr(respuesta) = 0 Then
'            MsgBox "Se ha pulsado cancelar o se ha cerrado el pregunta", vbInformation
'        ElseIf Len(respuesta) = 0 Then
'            MsgBox "Debe escribir un nombre", vbInformation
'        Else
'            strRespuesta = respuesta
'            NewPass = PonerPasswordVBA(CurrentDb.Name, "", strRespuesta)
'
'            If NewPass = False Then
'                MsgBox "No se ha podido cambiar la contraseña, porlo que deberás incluir un control de errores para entender porqué no se ha puesto"
'                Exit Sub
'            Else
'                MsgBox "Se ha protegido el proyecto con la contraseña '" & strRespuesta & "'. Recuerda guárdarla en un sitio seguro. Pulsa para cerrar el programa."
'                DoCmd.Quit
'            End If
'        End If
'    End If
'
'End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------------
  
      WizHook.Key = 51488399
      
      PonerPasswordVBA = WizHook.SetVBAPassword(strDBName, strConnect, strVBAPWD)

End Function
