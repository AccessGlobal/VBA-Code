Function OrdenaForm(frm As Form, ByVal sOrden As String, ByVal tipo As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-una-funcion-para-ordenar-todos-los-formularios
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : OrdenaForm
' Autor original    : desconocido
' Adaptado          : Luis Viadel | https://cowtechnologies.net
' Creado            : No lo recuerdo
' Propósito         : ordenar cualquier formulario por cualquier campo 
' Retorno           : verdadero/faso según se porduzca la ordenación o no
' Argumento/s       : La sintaxis de la función consta del siguiente argumento:
'                     Parte          Modo             Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm         Obligatorio      formulario que queremos ordenar
'                     sOrden      Obligatorio      campo por el que queremos ordenar
'                     tipo        Obligatorio      "ASC" o "DESC" según queramos ordenar ascendente o descendente
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : https://docs.microsoft.com/es-es/office/vba/api/Access.Form.OrderBy
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
' Sub OrdenaForm_test()
'
'        Call OrdenaForm(Me, "MiCampo", "ASC")
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
    Dim sform As String

    On Error GoTo LinErr
    
    OrdenaForm = False
    
    sform = frm.Name
    
    sOrden = sOrden & " " & tipo

    If frm.OrderByOn And (frm.OrderBy = sOrden) Then Exit Function
    
    frm.OrderBy = sOrden
    frm.OrderByOn = True
    
    OrdenaForm = True
    
    Exit Function

LinErr:
    OrdenaForm = False
    
End Function