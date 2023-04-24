Option Compare Database
Option Explicit
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/msgbox-personalizado/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : clsMsgBoxPersonalizado
' Autor original    : Luis Viadel | luisviadel@cowtechnologies.net
' Propósito         : crear una clase para manejar un mensaje personalizado para mostrar todos los mensajes de la aplicación
' Mas información   : Partiendo del formulario MsgBoxPersonalizado y sus objetos, se establecen 4 tipos diferentes de mensaje y 3 tipos de
'                     combinaciones de botones.
'                     Tipos:
'                            1. Exclamation
'                            2. Critical
'                            3. Question
'                            4. Information
'                     Botones:
'                            1. Aceptar / Cancelar
'                            2. Sí/no
'                            3. Aceptar
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega un módulo estándar. Descomenta la línea que nos interese y pulsa F5 para ver su funcionamiento.
'
'Sub clsMsgBoxPersonalizado_test()
'Dim cPerson As New clsMsgBoxPersonalizado
'Dim msgCabecera As String, msgTxt As String
'
'    msgCabecera = "Cabecera de mensaje"
'    msgTxt = "Este es un mensaje personalizado"
'
'    With cPerson
'        .Initialize 2, 3, msgCabecera, msgTxt
'    End With
'
'End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private ObjCab As Access.Label
Private Objmsg As Access.TextBox
Private ObjCuadro01 As Access.Rectangle
Private ObjCuadro02 As Access.Rectangle
Private ObjBtn01 As Access.CommandButton
Private ObjBtn02 As Access.CommandButton
Private ObjIcon As Access.image

Public Sub Initialize(msgTipo As Integer, msgButton As Integer, msgCabecera As String, msgTxt As String)

    DoCmd.OpenForm "MsgBoxPersonalizado", acNormal, , , , acHidden

'Cabecera
    Set ObjCab = Form_MsgBoxPersonalizado.msgCabecera
        With ObjCab
            .BackColor = -1 'Transparente
            .ForeColor = vbWhite
            .Caption = msgCabecera
        End With
    Set ObjCab = Nothing

'Texto del mensaje
    Set Objmsg = Form_MsgBoxPersonalizado.msgTxt
        With Objmsg
            .value = msgTxt
            .BackColor = -1
            .ForeColor = RGB(40, 40, 40)
'            .FontSize = 10
        End With
    Set Objmsg = Nothing

'Cuadro de cabecera
    Set ObjCuadro01 = Form_MsgBoxPersonalizado.msgCuadro01
        With ObjCuadro01
            Select Case msgTipo
                Case 1 'Exclamation
                    .BorderColor = RGB(231, 164, 0)
                    .BackColor = RGB(231, 164, 0)
                
                Case 2 'Critical
                    .BorderColor = RGB(193, 21, 25)
                    .BackColor = RGB(193, 21, 25)
                
                Case 3 'Question
                    .BorderColor = RGB(84, 130, 189)
                    .BackColor = RGB(84, 130, 189)
                
                Case 4 'Information
                    .BorderColor = RGB(73, 114, 215)
                    .BackColor = RGB(73, 114, 215)
            End Select
        End With
    Set ObjCuadro01 = Nothing

'Cuadro del cuerpo
    Set ObjCuadro02 = Form_MsgBoxPersonalizado.msgCuadro02
        With ObjCuadro02
            Select Case msgTipo
                Case 1 'Exclamation
                    .BorderColor = RGB(231, 164, 0)
                
                Case 2 'Critical
                    .BorderColor = RGB(193, 21, 25)
                
                Case 3 'Question
                    .BorderColor = RGB(84, 130, 189)
                
                Case 4 'Information
                    .BorderColor = RGB(73, 114, 215)
            End Select
        End With
    Set ObjCuadro02 = Nothing

'Botón 01
    Set ObjBtn01 = Form_MsgBoxPersonalizado.btn01
        With ObjBtn01
            Select Case msgButton
                Case 1 'Aceptar
                    .BackColor = RGB(193, 21, 25)
                    .BorderColor = RGB(193, 21, 25)
                    .PressedColor = RGB(186, 20, 25)
                    .PressedForeColor = RGB(193, 15, 18)
                    .FontSize = 9
                    .Caption = "Aceptar"
                    .Width = 960
                Case 2
                    .Caption = "Sí"
                    .Width = 540
                    .Move Left:=(Form_MsgBoxPersonalizado.btn01.Left + 800)
                Case 3
                    Form_MsgBoxPersonalizado.btn02.SetFocus
                    .Visible = False
            End Select
        End With
    Set ObjBtn01 = Nothing

'Botón 02
    Set ObjBtn02 = Form_MsgBoxPersonalizado.btn02
        With ObjBtn02
            Select Case msgButton
                Case 1
                    .FontSize = 9
                    .Caption = "Cancelar"
                    .Width = 960
                Case 2
                    .Caption = "No"
                    .Move Left:=(Form_MsgBoxPersonalizado.btn02.Left + 400)
                    .Width = 540
                Case 3
                    .BackColor = RGB(60, 102, 155)
                    .BorderColor = RGB(60, 102, 155)
                    .PressedColor = RGB(134, 167, 208)
                    .PressedForeColor = RGB(40, 68, 124)
            End Select
        End With
    Set ObjBtn02 = Nothing

'Icono del mensaje
    Set ObjIcon = Form_MsgBoxPersonalizado.msgIcon
        With ObjIcon
            Select Case msgTipo
                Case 1
                    .Picture = Application.CurrentProject.Path & "Exclamation.png"
                Case 2
                    .Picture = Application.CurrentProject.Path & "Critical.png"
                Case 3
                    .Picture = Application.CurrentProject.Path & "Question.png"
                Case 4
                    .Picture = Application.CurrentProject.Path & "Information.png"
            End Select
        End With
    Set ObjIcon = Nothing

'Mostramos el formulario
    Form_MsgBoxPersonalizado.Visible = True

End Sub
