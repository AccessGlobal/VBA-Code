'Colocar el siguiente código en los eventos de un formulario que contiene un listbox y un botón
Private Sub cmdCargarLista_Click()

     Call ListaComplementosAccess
    
End Sub

Private Sub lstLista_DblClick(Cancel As Integer)
Dim objAddin As clsAddIn
    
    Set objAddin = colAddins(Me.lstLista)
    
        MsgBox "Complemento:" & vbCrLf & objAddin.addin_Name & _
                vbNewLine & _
                "Librería:" & vbCrLf & _
                objAddin.library
                
    Set objAddin = Nothing
    
End Sub

'Colocar el siguiente código en un módulo estándar
Option Compare Database
Option Explicit

Const intForReading = 1
Const intUnicode = -1

Public colAddins As Collection

Private objFSO As Object

Sub ListaComplementosAccess()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-complementos-de-access
'                     Destello formativo 271
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ListaComplementosAccess
' Autor original    : Alba Salvá
' Creado            : 21/02/2023
' Adaptado por      : Luis Viadel
' Propósito         : listar todos los complementos de Access en un listbox de un formulario
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objAddin As clsAddIn
    
    DoCmd.Hourglass True
    
    Form_frmAddInsAccess.lstLista.RowSource = ""
    
    DoEvents
'Exportamos el registro a un fichero y lo recorremos creando una colección con los AddIns de Access
    ExportaRegistro
'Con la colección ya creada cargamos el listbox con sus elementos
    For Each objAddin In colAddins
        Form_frmAddInsAccess.lstLista.AddItem objAddin.addin_Name
    Next
    
    DoCmd.Hourglass False
    
    Set objAddin = Nothing
    
End Sub

Sub ExportaRegistro()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-complementos-de-access
'                     Destello formativo 271
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : ExportaRegistro
' Autor original    : Alba Salvá
' Creado            : 21/02/2023
' Adaptado por      : Luis Viadel
' Propósito         : exportar el registro de Windows a un fichero legible y manejable
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objShell As Object
Dim strRegPath As String
Dim strFileName As String, strCommand As String
Dim objRegFile As Object
Dim objInputFile As Object

    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
        strRegPath = "HKEY_LOCAL_MACHINE\Software\Microsoft\Office"
    
        strFileName = CurrentProject.Path & "\" & "Salida.reg"
    
        DoEvents
    
        strCommand = "cmd /c REG EXPORT " & strRegPath & " " & Replace(strRegPath, "\", "_") & ".reg /reg:64"
    
        objShell.Run strCommand, 0, True
    
        Set objRegFile = objFSO.CreateTextFile(strFileName, True, True)
            
            objRegFile.WriteLine "Windows Registry Editor Version 5.00"
            
            If objFSO.FileExists(Replace(strRegPath, "\", "_") & ".reg") = True Then
                Set objInputFile = objFSO.OpenTextFile(Replace(strRegPath, "\", "_") & ".reg", intForReading, False, intUnicode)
                    If Not objInputFile.AtEndOfStream Then
                        objInputFile.SkipLine
                        objRegFile.Write objInputFile.ReadAll
                    End If
                    objInputFile.Close
                Set objInputFile = Nothing
                
                objFSO.DeleteFile Replace(strRegPath, "\", "_") & ".reg", True
            End If
        
            objRegFile.Close
        Set objRegFile = Nothing
    
'Una vez que hemos exportado el fichero, la siguiente función lo lee buscando las claves de los complementos
        AnalizaRegistro strFileName
        
'Borrarmos el fichero en el que ha descargado el registro una vez que hemos encontrado lo que buscábamos
        objFSO.DeleteFile strFileName, True
    
    Set objShell = Nothing
    Set objFSO = Nothing
    
End Sub

Sub AnalizaRegistro(strFileName)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-complementos-de-access
'                     Destello formativo 271
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AnalizaRegistro
' Autor original    : Alba Salvá
' Creado            : 21/02/2023
' Adaptado por      : Luis Viadel
' Propósito         : lee el fichero de texto secuencialemnte en busca de los complementos.
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim objAddin As clsAddIn
Dim objInputFile As Object
Dim strLine As String
Dim salir As Boolean

    Set colAddins = Nothing
    
'Abrimos el fichero con los datos del registro y lo preparamos para lectura
    If objFSO.FileExists(strFileName) Then
        Set objInputFile = objFSO.OpenTextFile(strFileName, intForReading, False, intUnicode)
'Recorremos las líneas del fichero
        While Not objInputFile.AtEndOfStream
            
            strLine = objInputFile.readline
'Buscamos la cadena que nos va a mostrar los complementos "Access\Menu Add-Ins\"
            If InStr(strLine, "Access\Menu Add-Ins\") Then
                strLine = Mid(strLine, InStrRev(strLine, "\") + 1)
                strLine = Left(strLine, Len(strLine) - 1)
'Creamos una colección para los AddIns que vayamos encontrando
                If colAddins Is Nothing Then Set colAddins = New Collection
'Cada AddIn lo creamos mediante una clase que contiene el nombre y su librería
                Set objAddin = New clsAddIn
                
                objAddin.addin_Name = strLine
'Utilizamos una variable para crear un bucle que controlamos nosotros y que termina cuando lo decidimos
                salir = False
'Montamos el bucle
                While Not salir
                    DoEvents
                    strLine = objInputFile.readline
'Buscamos la cadena "Library" y lo añadimos a la clase
                    If InStr(strLine, "Library") Then
                        strLine = Mid(strLine, InStr(strLine, "=") + 1)
                        strLine = Replace(strLine, Chr(34), "")
                        strLine = Replace(strLine, "\\", "\")
                        objAddin.library = strLine
                        salir = True
                    End If
                Wend
'Lo añadimos a la colección
                colAddins.Add objAddin, objAddin.addin_Name
                
                Set objAddin = Nothing
            
            End If
'Cuando terminamos de recorrer las líneas del fichero del registro, finalizamos
        Wend
    End If
    
End Sub

'Colocar el siguiente código en un módulo de clase (clsAddIn)
Option Compare Database
Option Explicit

Public addin_Name As String
Public library As String
