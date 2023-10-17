Option Compare Database
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-numeros-aleatorios
'                     Destello formativo 362
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : NúmerosAleatorios
' Autor             : Luis Viadel | luisviadel@access-global.net
' Fecha             : octubre 223
' Propósito         : explorar las posibilidades de generación de números aleatorios en VBA
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://support.microsoft.com/en-us/office/rnd-function-503cd2e4-3949-413f-980a-ed8fb35c1d80
'                     https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/randomize-statement
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub btnGenerar_Click()

'Valor >0
'Genera aleatorios entre 0 y 1
'    Me.txtMonitor.Caption = Rnd(10)


'Valor <0
'Repite la secuencia
'     Me.txtMonitor.Caption = Rnd(-1)


'Valor =0
'Repite el más reciente
'     Me.txtMonitor.Caption = Rnd(0)


    Randomize
'Números aleatorios mayores que 1
'    Me.txtMonitor.Caption = Rnd * 100
    
'Número aleatorio entre 1 y 100
'    Me.txtMonitor.Caption = Int((100 * Rnd) + 1)
    
'Número aleatorio entre 100 y 1000
'    Me.txtMonitor.Caption = Int((1000 * Rnd) + 100)
    
    
'Número aleatorio entre 1 y 10 sin repetición
Dim matLista() As Long
Dim i As Long

    matLista = ListaUnica(10, 10, 1)
    
    For i = 1 To UBound(matLista)
        Debug.Print matLista(i)
    Next
    
End Sub

Function ListaUnica(NumValores As Long, numMax As Long, NumMin As Long) As Variant
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-numeros-aleatorios
'                     Destello formativo 362
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : NúmerosAleatorios
' Autor             : Luis Viadel | luisviadel@access-global.net
' Fecha             : octubre 223
' Propósito         : crear una lista de números aleatorios distintos
' Retorno           : matroz con la lista de números
' Argumento/s       : la sintaxis de la función consta de los siguientes argumentos:
'                     Parte              Modo           Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     NumValores      Obligatorio      número de valores aleatorios que queremos obtener
'                     numMax          Obligatorio      Valor máximo de la lista de valores
'                     numMin          Obligatorio      Valor mínimo de la lista de valores
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese, rellena los datos de la url y los del
'                     fichero que deseas descargar y pulsa F5 para ver su funcionamiento.
'
'Sub DescargaVersion_test()
'Dim matLista() As Long
'
'    matLista = ListaUnica(10, 10, 1)
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ListaRnd As Collection
Dim i As Long
Dim matLista() As Long
      
    Set ListaRnd = New Collection
    
        Randomize
     
        Do
            On Error Resume Next
            i = CLng(Rnd * (numMax - NumMin) + NumMin)
            
            ListaRnd.Add i, CStr(i)
                       
            On Error GoTo 0
        
        Loop Until ListaRnd.Count = NumValores
    
        ReDim matLista(1 To NumValores)
        
        For i = 1 To NumValores
            matLista(i) = ListaRnd(i)
        Next i
    
        ListaUnica = matLista()
        
    Set ListaRnd = Nothing
        
End Function
