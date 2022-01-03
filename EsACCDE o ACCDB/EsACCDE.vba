Public Function EsACCDE() As Boolean
    '--------------------------------------------------------------------------------------------------------
    ' Fuente            : https://access-global.net/es-accde-o-accdb/
    '--------------------------------------------------------------------------------------------------------
    ' Título            : EsACCDE
    ' Autor original    : Rafael Andrada | Mc. Pegasus
    ' Adaptado por      : Luis Viadel
    ' Actualizado       : Agosto 2020
    ' Propósito         : Saber si el fichero de Access es ACCDE o ACCDB
    ' Retorno           : devuelve true si el fichero es ACCDE y false ai es ACCDB
    '--------------------------------------------------------------------------------------------------------
    On Error GoTo LinError
    EsACCDE = (CurrentDb.Properties("MDE") = "T")
    ExitNow:
    Exit Function
    LinError:
    Resume ExitNow
    End Function