'Código en el formulario que queremos controlar
Private Sub FechaEjemplo_AfterUpdate()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            :  https://access-global.net/vba-abre-el-formulario-con-informacion-del-ultimo-acceso/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : FechaEjemplo_AfterUpdate
' Autor original    : Karl Donaubauer
' Fuente original   : https://accessusergroups.org/europe
' Adaptado          : Luis Viadel | https://cowtechnologies.net | luisviadel@cowtechnologies.net
' Creado            : desconocido
' Propósito         : utilizar las propiedades del textbox para guardar el último cambio realizado
' Referencia        : Microsoft XML, v6.0
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Más información   : En esta URL puedes descargar el fichero original
'                     https://accessusergroups.org/europe/wp-content/uploads/sites/22/2023/01/AEU_Tips.zip
'-----------------------------------------------------------------------------------------------------------------------------------------------


'Quasi-default value via DAO property

    On Error GoTo myError

    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim prp As DAO.Property

    Set db = CurrentDb
    
    'Form as DAO Document
    Set doc = db.Containers("Forms").Documents(Me.Name)
    
    'pass value to the property
    If Not IsNull(Me!FechaEjemplo) Then
        doc.Properties!prpDefaultDate = Me!FechaEjemplo
    End If


myExit:
    Exit Sub

myError:
    Select Case Err.Number
        Case 3270
            'property does not exist yet
            Set prp = doc.CreateProperty("prpDefaultDate", dbDate, Me!FechaEjemplo)
            doc.Properties.Append prp
            Resume Next
        Case Else
            MsgBox "Exception No. " & Err.Number & ". " & Err.Description
            Resume myExit
            Resume
    End Select

End Sub

Private Sub TxtEjemplo_AfterUpdate()

    On Error GoTo myError

    Dim db As DAO.Database
    Dim doc As DAO.Document
    Dim prp As DAO.Property

    If IsNull(Me.txtEjemplo) Then Me.txtEjemplo = "Escribe tu mensaje"
    
    Set db = CurrentDb
    
    'Form as DAO Document
    Set doc = db.Containers("Forms").Documents(Me.Name)
    
    'pass value to the property
    If Not IsNull(Me!txtEjemplo) Then
        doc.Properties!prpDefaultValue = Me!txtEjemplo
    End If


myExit:
    Exit Sub

myError:
    Select Case Err.Number
        Case 3270
            'property does not exist yet
            Set prp = doc.CreateProperty("prpDefaultValue", dbText, Me!txtEjemplo)
            doc.Properties.Append prp
            Resume Next
        Case Else
            MsgBox "Exception No. " & Err.Number & ". " & Err.Description
            Resume myExit
            Resume
    End Select

End Sub

Private Sub Form_Open(Cancel As Integer)

    On Error GoTo myError

    'pass value from property, 1-line notation because of lazin.... er example nature
    Me.FechaEjemplo = CurrentDb.Containers("Forms").Documents(Me.Name).Properties!prpDefaultDate
    Me.txtEjemplo = CurrentDb.Containers("Forms").Documents(Me.Name).Properties!prpDefaultValue
    

myExit:
    Exit Sub

myError:
    Select Case Err.Number
        Case 3270
            'property does not exist yet
            Resume myExit
        Case Else
            MsgBox "Exception No. " & Err.Number & ". " & Err.Description
            Resume myExit
            Resume
    End Select

    
End Sub
