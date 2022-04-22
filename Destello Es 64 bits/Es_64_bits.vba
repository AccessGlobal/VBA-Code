Public Function Es64Bit() As Boolean

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-es-accde
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Es64bits
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : junio 2010
' Propósito         : conocer si nuestra instalación de Oce es de 32 o de 64-bits
' Retorno           : booleano redadero/fals según sea o no de 64-bits
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencia        : https://docs.microsoft.com/es-es/office/vba/Language/Concepts/Getting-Started/64-bit-visual-basic-for-applications-overview
' Más informacion   : https://codekabinett.com/rdumps.php?Lang=2&targetDoc=windows-api-declaration-vba-64-bit
'-----------------------------------------------------------------------------------------------------------------------------------------------

#If Win64 Then
'Escribe aquí tú código para 64-bits
    Es64Bit = True
#Else
'Escribe aquí tu código para 32-bits
    Es64Bit = False
#End If

End Function