'----------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-kill-access
'----------------------------------------------------------------------------------------------------------
' Título            : KillAccess
' Autor original    : Desconocido
' Propósito         : Finalizar el proceso MSACCESS.
' Funcionamiento    : Crea un fichero de texto, pega este código en él y guárdalo como killaccess.vbs
'----------------------------------------------------------------------------------------------------------

strComputer = "."
strProcessToKill = "MSACCESS.exe" 

Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 
Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

For Each objProcess in colProcess
	objProcess.Terminate()
Next 

