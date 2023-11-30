Option Compare Database
Option Explicit

Public Const NumVersion = "01.01"
Public rstTable As DAO.Recordset

Public Sub GrabaEntrada()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-registro-de-accesos/
'                     Destello formativo 387
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GrabaEntrada
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 15/08/2013
' Propósito:        : Grabar en la base de datos los datos de inicio de sesión
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub GrabaEntrada_test()
'
'      Call GrabaEntrada
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
 Dim Objeto As Object
 Dim UsuName As String, CompName As String
 Dim HoraAcceso  As Date

    Set Objeto = CreateObject("WScript.Network")
'Tomamos el nombre del equipo
        CompName = Objeto.ComputerName
'Tomamos el usuario de Windows
        UsuName = Objeto.UserName
    Set Objeto = Nothing
                    
    HoraAcceso = Format(Time, "Long Time")
                    
    Set rstTable = CurrentDb.OpenRecordset("sesit")
        rstTable.AddNew
            rstTable!IdTRABA = 1 'Indicar el Id del usuario conectado
            rstTable!SESITFI = Now
            rstTable!SESITHI = HoraAcceso
            rstTable!SESITIN = 1 'Se puede añadir un contador de accesos que registre cada vez que falla en indicar la contraseña
            rstTable!SESITIP = GetMyLocalIP()
            rstTable!SESITIPE = GetMyPublicIP()
            rstTable!SESITM = CompName
            rstTable!SESITUS = UsuName
            rstTable!sesitver = NumVersion
        rstTable.Update
        rstTable.Close
    Set rstTable = Nothing

End Sub

Public Sub GrabaSalida()
'---------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-registro-de-accesos/
'                     Destello formativo 387
'---------------------------------------------------------------------------------------------------------------------------------------------
' Título            : GrabaSalida
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 15/08/2013
' Propósito:        : Grabar en la base de datos los datos de finalización de sesión
'---------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA. Pulsa F5 para ver su funcionamiento.
'
' Sub GrabaEntrada_test()
'
'      Call GrabaSalida
'
' End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim hf As Date
Dim Tiempo As Single
Dim strOutput As Long
Dim IdTRABA As Integer

    IdTRABA = 1 'Indicar el Id del usuario conectado
    
    hf = Format(Time, "Long Time")
    
'Busca el registro en el que no hemos indicado la salida, que es el registro actual
    Set rstTable = CurrentDb.OpenRecordset("SELECT * FROM SESIT WHERE IdTRABA=" & IdTRABA & " AND IsNull(SESITFF) ORDER BY IdSESIT DESC")
        Do Until rstTable.EOF
            rstTable.Edit
                Tiempo = Now - rstTable!SESITFI
'Convertimos el tiempo transcurrido en formato hh/mm/ss
                strOutput = Int(CSng(Tiempo * 24 * 3600))
                rstTable!usua = Format$(CInt(CInt(strOutput / 60) / 60), "00") & ":" & _
                                Format$(CStr(Int(strOutput / 60) Mod 60), "00") & ":" & _
                                Format$(CStr(strOutput Mod 60), "00")
                rstTable!sesittime = strOutput
                rstTable!SESITHF = hf
                rstTable!SESITFF = Now
            rstTable.Update
        rstTable.MoveNext
        Loop
    Set rstTable = Nothing

End Sub