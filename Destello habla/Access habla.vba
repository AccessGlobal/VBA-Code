Public Function AccessHabla(ByVal strFrase As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-que-te-lo-diga-access
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AccessHabla
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 22
' Propósito         : utilizamos la SAPI 5.3 (SpVoice interface) para que Access "lea" la información que deseamos
' Retorno           : lectura de la frase indicada, directamente por los altavoces
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms723602(v=vs.85)?redirectedfrom=MSDN
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub CheckInternet_test()
'
'        AccessHabla ("Recuerda desconectar la alarma y cerrar todas las ventanas")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objSpeech As Object
Dim tono As Integer

Set objSpeech = CreateObject("SAPI.SpVoice")
    tono = 1
    objSpeech.Speak "<pitch middle = '" & tono & "'/>" & strFrase 'Tono medio
Set objSpeech = Nothing

End Function
