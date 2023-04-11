Function FormularioAbierto(strNombreFormulario As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-esta-abierto-el-formulario/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FormularioAbierto
' Autor original    : Alba Salvá
' Creado            : Desconocido
' Propósito         : Saber, en tiempo de ejecución, si un formulario está abierto.
' Retorno           : Verdadero o Falso, según si está abierto (True) o cerrado (False)
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte                     Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strNombreFormulario       Obligatorio    Nombre del formulario que se desea consultar
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, rellena los datos de la url y los del
'                     fichero que deseas descargar y pulsa F5 para ver su funcionamiento.
'
'Sub FormularioAbierto_test()
'
'   FormularioAbierto "NombreFormulario"
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim i As Integer

    For i = 0 To Forms.Count - 1
        If Forms(i).Name = strNombreFormulario Then
            FormularioAbierto = True
            Exit Function
        End If
    Next

End Function
