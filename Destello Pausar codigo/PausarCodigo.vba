Option Compare Database
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-la-forma-mas-eficiente-de-pausar-el-codigo
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Diversas formas de hacer una pausa
' Propósito         : realizar una pausa controlada de nuestro código
' Retorno           : Sin retorno
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://docs.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-gettickcount
'                     https://access-global.net/hagamos-una-pausa/
'                     https://docs.microsoft.com/en-us/windows/win32/api/synchapi/nf-synchapi-sleep
'                     https://docs.microsoft.com/en-us/windows/win32/sync/wait-functions
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Sub Pausa01_test()
' Call EsperarConGetTick(3)
'End Sub
'
'Sub Pausa02_test
' sleep 3000
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Public Function EsperarConGetTick(nsegundos As Long)
Dim InicioTick As Long
Dim ActualTick As Long
Dim FinTick As Long

On Error GoTo LinErr:

InicioTick = GetTickCount
FinTick = GetTickCount + (nsegundos * 1000)

Do
    ActualTick = GetTickCount
    DoEvents
Loop Until (ActualTick >= FinTick)

Exit Function
    
    
LinErr:

End Function