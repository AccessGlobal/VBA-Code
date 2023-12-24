

'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-control-ficha/
'                     Destello formativo 394
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Control página, métodos para seleccionar una ficha, página o pestaña por código (Access VBA)
' Autor original    : Rafael .:McPegasus:.
' Creado            : 18/12/2023
' Propósito         : De un control de tipo página, conocer los diferentes métodos para seleccionar una ficha (página o pestaña) por código.
'---------------------------------------------------------------------------------------------------------------------------------------------
' Información       : https://learn.microsoft.com/es-es/office/vba/api/overview/tab-control
'-----------------------------------------------------------------------------------------------------------------------------------------------

Private Sub cmdAnteriorPestaña_Click()
Dim intPagesCount                               As Integer
Dim intPagePrevious                             As Integer
        
    intPagesCount = Me.tabEjemplo.Pages.Count
    intPagePrevious = Me.tabEjemplo.Value - 1
    intPagePrevious = IIf(intPagePrevious = -1, intPagesCount - 1, intPagePrevious)

    Me.tabEjemplo = intPagePrevious

End Sub

Private Sub cmdSiguientePestaña_Click()
Dim intPagesCount                               As Integer
Dim intPageNext                                 As Integer
        
    intPagesCount = Me.tabEjemplo.Pages.Count
    intPageNext = Me.tabEjemplo.Value + 1
    intPageNext = IIf(intPageNext = intPagesCount, 0, intPageNext)

    Me.tabEjemplo = intPageNext

End Sub

Private Sub cmdPrimeraPestaña_Click()

    'Hay tres métodos posibles, que son los siguientes.
    Me.tabEjemplo = 0
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item(0).PageIndex
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item("pgnPestaña01").PageIndex

End Sub

Private Sub cmdSegundaPestaña_Click()

    'Hay tres métodos posibles, que son los siguientes.
    Me.tabEjemplo = 1
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item(1).PageIndex
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item("pgnPestaña02").PageIndex

End Sub

Private Sub cmdTerceraPestaña_Click()

    'Hay tres métodos posibles, que son los siguientes.
    Me.tabEjemplo = 2
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item(2).PageIndex
    Me.tabEjemplo = Me.tabEjemplo.Pages.Item("pgnPestaña03").PageIndex
    
End Sub
