Public Function EstaAbierto(ByVal frm As String) As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-estado-de-un-formulario/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : EstaAbierto
' Autor original    : Luis Viadel
' Creado            : 19/08/2003
' Propósito         : Saber, en tiempo de ejecución, si un formulario está abierto.
' Retorno           : Verdadero o Falso, según si está abierto (True) o cerrado (False)
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte     Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     frm       Obligatorio    Nombre del formulario que se desea consultar
'
'-----------------------------------------------------------------------------------------------------------------------------------------------

EstaAbierto = SysCmd(acSysCmdGetObjectState, acForm, frm)

End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------