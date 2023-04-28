Public Function WH_WizMsgBox_test()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/Wizhook-series-WizMsgBox/
'                     Destello formativo 315
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : WH_WizMsgBox_test
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 2023
' Propósito         : mostrar un cuadro de mensaje
'------------------------------------------------------------------------------------------------------------------------------------
Dim resultado As Long
Dim strText As String
Dim strCaption As String
Dim strHFile As String

    strCaption = "Ejemplo de WizHook msgbox"
    strText = vbCrLf & _
              "Primera línea de ejemplo" & vbCrLf & _
              "Segunda línea de ejemplo." & _
               vbCrLf
'Si queremos incluir un fichero de ayuda
    strHFile = "C:\MiFichero.chm"
    
    WizHook.key = 51488399
    
    resultado = WizHook.WizMsgBox(strText, strCaption, vbCritical + vbOKCancel, 1, strHFile)
    
    Debug.Print resultado
    
End Function
