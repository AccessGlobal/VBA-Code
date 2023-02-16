'Añadir en el evento "Al hacer click" de un botón
Private Sub btnAddEvent_Click()
Dim vbc As VBIDE.VBComponent
Dim strMod As String, strObject As String
Dim strEvent As String, strCode As String
Dim strProc As String
Dim proctipo As String, intStr As Integer
Dim lngProcTyp As Long
Dim lngStartLine As Long
Dim i As Integer

'Si el usuario no selecciona ningún módulo sale
    If IsNull(Me.lstProc.Value) Then
        MsgBox "Debe seleccionar un módulo de la lista"
        Exit Sub
    End If
    
'Seleccionamos el módulo que el usuario ha seleccionado en el listbox
    strMod = Me.lstProc.Value
    
    intStr = InStr(1, strMod, " ")
'Extraemos el nombre
    If intStr = 0 Then
        strMod = strMod
    Else
        strMod = Left(strMod, intStr - 1)
    End If
    
'Comprobamos que es un módulo y no un procedimiento y si es un módulo
'borramos el procedimiento que queremos sustituir

    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        
        If strMod = vbc.Name Then GoTo LinContinua

    Next
    
    MsgBox "El elemento seleccionado no es un módulo"
    
    Exit Sub
    
LinContinua:
'Abrimos el formulario que lo contiene en oculto para que no intervenga el usuario
    DoCmd.OpenForm "ProcedimientosAddEventTest", acDesign, , , , acHidden ' Tenemos que abrirlo oculto

'Como el botón ya tiene código, borramos primero el procedimiento

    strProc = "btnTest_Click"
    
    Call BorraProcedimiento(strMod, strProc, vbext_pk_Proc)
    
' El objeto sobre el que vamos a escribir el evento es el botón btnAddEvent1
    strObject = "btnTest"
    
    strEvent = "Click"

    strCode = vbNewLine & "    msgbox" & """Este es el evento del botón del usuario 1"""
                
    Call AgregaEvento(strMod, strObject, strEvent, strCode)
    
    DoCmd.Close acForm, "ProcedimientosAddEventTest", acSaveYes
    
    MsgBox "Ya he añadido el evento del botón, púlsamo para comprobar"
    
    
End Sub

'Código a incluir en un módulo estándar
Option Compare Database
Option Explicit

Public Sub AgregaEvento(strModule As String, strObject As String, strEvent As String, strCode As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-series-procedimientos-cambiamos-el-evento-click-de-un-boton-en-tiempo-de-ejecucion/
'                     Destello formativo 267
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : AgregaEvento
' Autor original    : Alba Salvá
' Creado            : desconocido
' Adaptado por      : Luis Viadel
' Propósito         : Graba un nuevo evento en tiempo de ejecución
' Argumentos        : La sintaxis del procedimiento consta de los siguientes argumentos:
'                     Parte           Modo          Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     strModule    Obligatorio      Nombre del módulo que contiene el procedimiento
'                     strObject    Obligatorio      Objeto sobre el que queremos actuar
'                     strEvent     Obligatorio      Tipo de evento
'                     strCode      Obligatorio      Código que vamos a incluir en el evento
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : Microsoft Visual Basic for Applications Extensibility 5.3
'                     C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'                     {0002E157-0000-0000-C000-000000000046}
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente en el
'                     evento de un botón. No olvides dar valor a las variables strModuloName (Nombre del módulo donde vas a crear el
'                     procedimiento) y strVBACode (que contiene el procedimiento)
'                     Descomenta la línea que te interese y pulsa F5 para ver su funcionamiento.
'
'Private Sub MiBoton_Click()
'
'    Call AgregaEvento("Nombre del módulo", "Nombre del objeto", "Tipo de evento", "Código VBA que añadimos")
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------Dim lngStartLine As String
Dim lngStartLine As Long

    With Application.VBE.ActiveVBProject.VBComponents(strModule).CodeModule
        lngStartLine = .CreateEventProc(strEvent, strObject) + 1
        .InsertLines lngStartLine, strCode
    End With

End Sub