Public Sub WH_CerrarBD_test()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/wizhook-series-3-metodos-para-cerrar-una-base-de-datos-y-wizhook/
'                     Destello formativo 316
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : WH_CerrarBD_test
' Autor original    : Luis Viadel | https://cowtechnologies.net
' Creado            : abril 2023
' Propósito         : diferentes formas de cerrar nuestra base de datos
'------------------------------------------------------------------------------------------------------------------------------------
    
'Método 1
'Objeto Docmd

'    DoCmd.Quit acQuitSaveAll

'Método 2
'Objeto Application

'    Application.Quit

'Método 3
'Objeto Application II
    
'    Application.CloseCurrentDatabase

'Método 4
'WizHook
    
    WizHook.key = 51488399
    
    Debug.Print WizHook.CloseCurrentDatabase()
       
       
End Sub
