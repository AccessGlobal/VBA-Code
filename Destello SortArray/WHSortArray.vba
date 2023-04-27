Public Function WH_sort_test()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-/
'                     Destello formativo 314
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : WH_sort_test
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 2010
' Propósito         : ordenar alfábeticamente los elementos de una matriz
'------------------------------------------------------------------------------------------------------------------------------------
Dim mat(5) As String
Dim cont As Integer

    mat(0) = "Naranja"
    mat(1) = "Zanahoria"
    mat(2) = "Aguacate"
    mat(3) = "Pera"
    mat(4) = "Piña"
    mat(5) = "Plátano"
    
    WizHook.SortStringArray mat
    
    For cont = 0 To 5
        Debug.Print mat(cont)
    Next cont
    

End Function