Public Sub mcOLK_Establecer_SendUsingAccount_AUGE_Happy()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/establecer-una-cuenta-personalizada-en-un-nuevo-mensaje-de-correo-de-outlook/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : mcOLK_Establecer_SendUsingAccount_AUGE_Happy
' Autor             : Rafael Andrada .:McPegasus:. | BeeSoftware
' Actualizado       : 21/01/2022
' Propósito         : Establecer en un nuevo correo utilizando Micrsofot Outlook la cuenta de correo del remitente, una diferente a la predeterminada.
' Retorno           : No hay retorno por ser un procedimiento pero en caso de no producirse ningún error nos encontramos con un mensaje abierto y con la cuenta de envío (De:) según establecida en una variable.
' Sobre Referenciar : El referenciar una librería externa nos permite seleccionar los objetos de otra aplicación que se desea que estén disponibles en nuestro código.
'                     También acceder a sus métodos utilizar las constantes.
'                     En caso de ser opcional podemos seguir utilizándolo aunque las constantes hay que sustituirlas por su valor, normalmente numérico.
'                     Más información: https://support.microsoft.com/es-es/office/add-object-libraries-to-your-visual-basic-project-ed28a713-5401-41b0-90ed-b368f9ae2513
' Referencia        : Opcional o no según el valor de cblnSeUsaReferencia_OL_AUGE_Happy. Microsoft Outlook 16.0 Object Library (c:\Program Files (x86)\Microsoft Office\root\Office16\MSOUTL.OLB)
' Importante        : Dedico este módulo a Happy por resolver en el Whatsapp de Access User Group España un problema en la línea que se indica en el código.
' Importante        : En este ejemplo se presuponde los siguientes puntos _
                            1 - Que MS Outlook está instalado en el equipo que va a ejecutar el test. _
                            2 - Que la cuenta que establezcamos en la variable strDe esté configurada en el equipo, en nuestro MS Outlook. Lo ideal es tener más de una  configurada para cambiar y comprobar el proceso./>
'-----------------------------------------------------------------------------------------------------------------------------------------------
   
    'Microsoft Outlook 16.0 Object Library. (c:\Program Files (x86)\Microsoft Office\root\Office16\MSOUTL.OLB)
    #Const cblnSeUsaReferencia_OL_AUGE_Happy = False

     #If cblnSeUsaReferencia_OL_AUGE_Happy Then
         Dim olkApp                                 As Outlook.Application
         Dim olkMailItem                            As Outlook.MailItem

         Dim olkAccountUser                         As Outlook.Account
         Dim olkAccountsUsers                       As Outlook.Accounts

     #Else
         Dim olkApp                                 As Object
         Dim olkMailItem                            As Object

         Dim olkAccountUser                         As Object
         Dim olkAccountsUsers                       As Object

     #End If

    Dim strDe                                       As String
    
    
    Set olkApp = CreateObject("Outlook.Application")
    
    Set olkMailItem = olkApp.CreateItem(0)

'*********************************************************************************************************************************************************
' OJO hay que sustituir por una cuenta que esté configurada en nuestro MS OL, en caso contrario se producirá un error que no es de mi interés capturarlo en este módulo.
    strDe = "rafael@mcpegasus.net"
    strDe = "rafael.andradas@access-global.net"
'*********************************************************************************************************************************************************
    
    '21/01/2022 09:55 Se produce un error en la siguiente línea por lo que me obliga a declarar los objetos referenciados a Outlook en este procedimiento.
    '21/01/2022 10:04 Al probar con declaración As Object es cuando se produce el error.
    '21/01/2022 11:13 El problema viene al declarar la variable olkApp como As Object. En el caso de Outlook.Application funciona correctamente.
    '21/01/2022 11:47 Pregunta en AUGE por si alguien sabe tocar la flauta.
'*********************************************************************************************************************************************************
    'En la siguiente línea se produce el error 450: El número de argumentos es incorrecto o la asignación de propiedad no es válida.
'    Set olkAccountUser = olkApp.Session.Accounts(strDe)
'*********************************************************************************************************************************************************
    '21/01/2022 13:00 Y Happy tocó la flauta solucionándolo con las siguientes líneas.
    
    #If cblnSeUsaReferencia_OL_AUGE_Happy Then
        Set olkAccountUser = olkApp.Session.Accounts(strDe)
        
    #Else
        Set olkAccountsUsers = olkApp.Session.Accounts
        Set olkAccountUser = olkAccountsUsers(strDe)
        
    #End If

    Set olkMailItem = olkApp.CreateItem(0)

    #If cblnSeUsaReferencia_OL_AUGE_Happy Then
        olkMailItem.SendUsingAccount = olkAccountUser
        
    #Else
        Set olkMailItem.SendUsingAccount = olkAccountUser
        
    #End If
    
    'Otra solución que se propuso y que tambíen funciona es hacerlo utilizando un For Each.
'    #If cblnSeUsaReferencia_OL_AUGE_Happy Then
'        Set olkAccountUser = olkApp.Session.Accounts(strDe)
'        Set olkMailItem = olkApp.CreateItem(0)
'        olkMailItem.SendUsingAccount = olkAccountUser
'
'    #Else
'        For Each olkAccountUser In olkAccountsUsers
'            If olkAccountUser = strDe Then
'                Set olkMailItem.SendUsingAccount = olkAccountUser
'
'            End If
'        Next
'    #End If

    olkMailItem.Display                         'Mostrar el mensaje saliente en pantalla.

    If Not olkAccountUser Is Nothing Then Set olkAccountUser = Nothing
    If Not olkAccountsUsers Is Nothing Then Set olkAccountsUsers = Nothing
    If Not olkApp Is Nothing Then Set olkApp = Nothing

End Sub
