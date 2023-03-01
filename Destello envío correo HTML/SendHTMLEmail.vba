'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/diseno-enviar-bonitos-correos-en-html/
'-----------------------------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : SendHTMLEmail
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Automate Outlook to send an HTML email with or without attachments
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Late Binding version  -> None required
'             Early Binding version -> Ref to Microsoft Outlook XX.X Object Library
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sTo       : To Recipient email address string (semi-colon separated list)
' sSubject  : Text string (HTML) to be used as the email subject line
' sBody     : Text string to be used as the email body (actual message)
' bEdit     : True/False whether or not you wish to preview the email before sending
' vCC       : CC Recipient email address string (semi-colon separated list)
' vBCC      : BCC Recipient email address string (semi-colon separated list)
' vAttachments : Array of attachment (complete file paths with
'                   filename and extensions)
' vAccount  : Name of the Account to use for sending the email (normally the e-mail adddress)
'                   if no match is found it uses the default account
'
' Usage:
' ~~~~~~
' Call SendHTMLEmail("abc@xyz.com", "My Subject", "<p>My <b>body</b>.</p>", True)
' Call SendHTMLEmail("abc@xyz.com;def@wuv.ca;", "My Subject", "<p>My <b>body</b>.</p>", True)
' Call SendHTMLEmail("abc@xyz.com", "My Subject", "<p>My <b>body</b>.</p>", True, , _
'                    Array("C:\Temp\Table2.txt"))
' Call SendHTMLEmail("abc@xyz.com", "My Subject", "<p>My <b>body</b>.</p>", True, , _
'                    Array("C:\Temp\Table2.txt", "C:\Temp\Supplier List.txt"))
' Call SendHTMLEmail("abc@xyz.com", "My Subject", "<p>My <b>body</b>.</p>", True, , _
'                    Array("C:\Temp\Table2.txt", "C:\Temp\Supplier List.txt"), _
'                    "cde@uvw.com")
' Call SendHTMLEmail("abc@xyz.com", "My Subject", "<p>My <b>body</b>.</p>", True, , _
'                    Split("C:\Temp\Table2.txt,C:\Temp\Supplier List.txt", ","), _
'                    "cde@uvw.com")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2007-11-16              Initial Release
' 2         2017-02-15              Added retention of default e-mail signature
'                                   Added conditional compiler directives for early and
'                                       late binding
' 3         2019-01-20              Updated Copyright
'                                   Added usage examples
'                                   Added sAccount option
' 4         2019-09-06              Updated the handling of sTo and sBCC to split e-mail
'                                       addresses into individual recipients and
'                                       improved error reporting for unresolvable e-mail
'                                       addresses per an issue flagged by InnVis (MSDN)
' 5         2020-03-12              Bugs fixes (missing declarations) from comments by
'                                       S.A.Marshall in answers forum
'                                   Added CC to function
' 6         2020-12-08              Proper parsing of .HTMLBody so proper HTML is
'                                       generated
' 7         2021-02-19              Added Split() example to usage examples
'---------------------------------------------------------------------------------------
Function SendHTMLEmail(ByVal sTo As String, _
                        ByVal sSubject As String, _
                        ByVal sBody As String, _
                        ByVal bEdit As Boolean, _
                        Optional vCC As Variant, _
                        Optional vBCC As Variant, _
                        Optional vAttachments As Variant, _
                        Optional vAccount As Variant) As Boolean
    On Error GoTo Error_Handler
    '    #Const EarlyBind = 1 'Use Early Binding
    #Const EarlyBind = 0    'Use Late Binding
    #If EarlyBind Then
        Dim oOutlook          As Outlook.Application
        Dim oOutlookMsg       As Outlook.MailItem
        Dim oOutlookInsp      As Outlook.Inspector
        Dim oOutlookRecip     As Outlook.Recipient
        Dim oOutlookAttach    As Outlook.Attachment
        Dim oOutlookAccount   As Outlook.Account
    #Else
        Dim oOutlook          As Object
        Dim oOutlookMsg       As Object
        Dim oOutlookInsp      As Object
        Dim oOutlookRecip     As Object
        Dim oOutlookAttach    As Object
        Dim oOutlookAccount   As Object
        Const olMailItem = 0
    #End If
    Dim sHTML                 As String
    Dim aHTML                 As Variant
    Dim aSubHTML              As Variant
    Dim aRecip                As Variant
    Dim I                     As Integer
 
    Set oOutlook = CreateObject("Outlook.Application")
    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
 
    With oOutlookMsg
        'Account to use for sending, if specified, otherwise use default
        If Not IsMissing(vAccount) Then
            For Each oOutlookAccount In oOutlook.Session.Accounts
                If oOutlookAccount = vAccount Then
                    Set oOutlookMsg.SendUsingAccount = oOutlookAccount
                End If
            Next
        End If
    DoCmd.SetWarnings False
'        .display    'Had to move this command here to resolve a bug only existent in Access 2016!
 
        'TO
        aRecip = Split(sTo, ";")
        For I = 0 To UBound(aRecip)
            If Trim(aRecip(I) & "") <> "" Then
                Set oOutlookRecip = .Recipients.Add(aRecip(I))
                oOutlookRecip.Type = 1
            End If
        Next I
'        .display    'Had to move this command here to resolve a bug only existent in Access 2016!
        'CC
        If Not IsMissing(vCC) Then
            aRecip = Split(vCC, ";")
            For I = 0 To UBound(aRecip)
                If Trim(aRecip(I) & "") <> "" Then
                    Set oOutlookRecip = .Recipients.Add(aRecip(I))
                    oOutlookRecip.Type = 2
                End If
            Next I
        End If
 
        'BCC
        If Not IsMissing(vBCC) Then
            aRecip = Split(vBCC, ";")
            For I = 0 To UBound(aRecip)
                If Trim(aRecip(I) & "") <> "" Then
                    Set oOutlookRecip = .Recipients.Add(aRecip(I))
                    oOutlookRecip.Type = 3
                End If
            Next I
        End If
 
        .subject = sSubject
'        Set oOutlookInsp = .GetInspector    'Retains the signature if applicable
 
        sHTML = .HTMLBody
        aHTML = Split(sHTML, "<body")
        aSubHTML = Split(aHTML(1), ">")
        sHTML = aHTML(0) & "<body" & aSubHTML(0) & ">" & _
                sBody & _
                Right(aHTML(1), Len(aHTML(1)) - Len(aSubHTML(0) & ">"))
        .HTMLBody = sHTML
        .Importance = 2    'Importance Level  0=Low,1=Normal,2=High
 
        ' Add attachments to the message.
        If Not IsMissing(vAttachments) Then
            If IsArray(vAttachments) Then
                For I = LBound(vAttachments) To UBound(vAttachments)
                    If vAttachments(I) <> "" And vAttachments(I) <> "False" Then
                        Set oOutlookAttach = .Attachments.Add(vAttachments(I))
                    End If
                Next I
            Else
                If vAttachments <> "" Then
                    Set oOutlookAttach = .Attachments.Add(vAttachments)
                End If
            End If
        End If
 
        For Each oOutlookRecip In .Recipients
            If Not oOutlookRecip.Resolve Then
                'You may wish to make this a MsgBox! to show the user that there is a problem
                Debug.Print "Could not resolve the e-mail address: ", oOutlookRecip.Name, oOutlookRecip.Address, _
                            Switch(oOutlookRecip.Type = 1, "TO", _
                                   oOutlookRecip.Type = 2, "CC", _
                                   oOutlookRecip.Type = 3, "BCC")
                bEdit = True    'Problem so let display the message to the user so they can address it.
            End If
        Next
 
'        If bEdit = True Then    'Choose btw transparent/silent send and preview send
'            '.Display
'        Else
            .Send
            SendHTMLEmail = True

'        End If
    End With
 
Error_Handler_Exit:
    On Error Resume Next
    If Not oOutlookAccount Is Nothing Then Set oOutlookAccount = Nothing
    If Not oOutlookAttach Is Nothing Then Set oOutlookAttach = Nothing
    If Not oOutlookRecip Is Nothing Then Set oOutlookRecip = Nothing
    If Not oOutlookInsp Is Nothing Then Set oOutlookInsp = Nothing
    If Not oOutlookMsg Is Nothing Then Set oOutlookMsg = Nothing
    If Not oOutlook Is Nothing Then Set oOutlook = Nothing
    Exit Function
 
Error_Handler:
    SendHTMLEmail = False

    If Err.Number = "287" Then
        MsgBox "You clicked No to the Outlook security warning. " & _
               "Rerun the procedure and click Yes to access e-mail " & _
               "addresses to send your message. For more information, " & _
               "see the document at http://www.microsoft.com/office" & _
               "/previous/outlook/downloads/security.asp."
    Else
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Source: SendHTMLEmail" & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
End Function

