Option Compare Database
Option Explicit

Public Function AbrirEstadisticas()

    DoCmd.OpenForm "VBIDEEstadisticas"
    
End Function


Public Function VBIDE_Estadisticas() As Variant
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vbide-algunas-estadisticas/
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Título            : EstadisticasReferencias
' Autor original    : Luis Viadel
' Creado            : desconocido
' Colaboradores     : Alba Salvá 
' Propósito         : Obtener cirtos valores de propiedades del VBIDE del programa host, ya que se realiza la acción desde un complemento
' Argumento/s       : no tiene argumentos
' Devolución        : la función devuelve un matriz con todos los valores que queremos calcular
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       :
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Test:             : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test.  Copia el bloque siguiente código al
'                     portapapeles y pega en el editor de VBA, en el evento del objeto que desees. En el test lo hacemos desde la carga de un form
'                     Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'Private Sub Form_Load()
'Dim VBIDEStats As Variant
'
'    VBIDEStats = VBIDE_Estadisticas
'
'End Sub
'
'---------------------------------------------------------------------------------------------------------------------------------------------------
Dim objRef As VBIDE.Reference
Dim objRefs As VBIDE.References
Dim RefNum As Integer
Dim vbc As VBIDE.VBComponent
Dim vbcProjects As VBIDE.VBProjects
Dim vbcProject As VBIDE.VBProject
Dim NumStats As Variant
Dim NumMods As Integer, NumMod1 As Integer, NumMod2 As Integer
Dim NumMod3 As Integer, NumMod4 As Integer, NumMod5 As Integer
Dim NumProcs As Integer
Dim NumLinProcs As Long
Dim NumLinProcsCabecera As Long
Dim lngLastLine As Long, lngStartLine As Long
Dim lngLastLine1 As Long, lngStartLine1 As Long
Dim lngPublic As Long, lngPrivate As Long, lngComment As Long
Dim lngSub As Long, lngFunction As Long, lngVariables As Long

'Calculamos el número de referencias
    Set objRefs = Application.VBE.ActiveVBProject.References
        For Each objRef In objRefs
            RefNum = RefNum + 1
        Next
    Set objRef = Nothing

'Leyenda para módulos
'NumMod1=Número de módulos estándar
'NumMod2=Número de módulos de clase
'NumMod3=Número de UserForms
'NumMod4=Número de Formularios
'Nummod5=Número de informes

'Inicializamos todas las variables
    NumMod1 = 0
    NumMod2 = 0
    NumMod3 = 0
    NumMod4 = 0
    NumMod5 = 0

'Localizamos el VBProject correcto, porque si no, nos mostrará siempre el del complemento
    Set vbcProjects = Application.VBE.VBProjects
        For Each vbcProject In vbcProjects
            If vbcProject.Name <> "VBIDE_Estadisticas" Then
                If vbcProject.Name <> "ACWZTOOL" Then
'Recorremos los módulos y vamos revisando sus tipos
                    For Each vbc In vbcProject.VBComponents
                        Debug.Print vbc.Name
                        NumMods = NumMods + 1
                        Select Case vbc.Type
                            Case 1
                                NumMod1 = NumMod1 + 1
                    
                            Case 2
                                NumMod2 = NumMod2 + 1
                    
                            Case 3
                                NumMod3 = NumMod3 + 1
                    
                            Case 100
                                
                                If Left(vbc.Name, 4) = "Form" Then
                                    NumMod4 = NumMod4 + 1
                                Else
                                    NumMod5 = NumMod5 + 1
                                End If
                        End Select
                    Next
                    Exit For
                End If
            End If
        Next
    Set vbcProjects = Nothing
    
'Repetimos la operación para los procedimientos
'Inicializamos las variables
        NumLinProcsCabecera = 0
        NumLinProcs = 0
        NumProcs = 0
        lngPublic = 0
        lngPrivate = 0
        lngComment = 0
        lngSub = 0
        lngFunction = 0
        lngLastLine = 0
        lngStartLine = 0
    
'Recorremos todos los módulos
    Set vbcProjects = Application.VBE.VBProjects
        For Each vbcProject In vbcProjects
            If vbcProject.Name <> "VBIDE_Estadísticas" Then
                If vbcProject.Name <> "ACWZTOOL" Then
                    For Each vbc In vbcProject.VBComponents
                        With vbc.CodeModule
                            lngStartLine = .CountOfDeclarationLines + 1
                            lngLastLine = .CountOfLines
                            lngLastLine1 = .CountOfLines
                            lngStartLine1 = .CountOfDeclarationLines + 1
                            NumLinProcsCabecera = NumLinProcsCabecera + .CountOfDeclarationLines
                            NumLinProcs = NumLinProcs + .CountOfLines - .CountOfDeclarationLines
            
'Contamos los procedimientos
                            Do Until lngStartLine >= .CountOfLines
                                lngStartLine = lngStartLine + .ProcCountLines(.ProcOfLine(lngStartLine, vbext_pk_Proc), vbext_pk_Proc)
                                NumProcs = NumProcs + 1
                            Loop
            
'Buscamos concepto a concepto
                            For lngStartLine = lngStartLine1 To lngLastLine1
                                If vbc.CodeModule.Find("Public Sub", lngStartLine, 1, lngStartLine, 10) = True Then
                                    lngSub = lngSub + 1
                                End If
                            Next
                                        
                            For lngStartLine = lngStartLine1 To lngLastLine1
                                If vbc.CodeModule.Find("Private Sub", lngStartLine, 1, lngStartLine, 11) = True Then
                                    lngSub = lngSub + 1
                                    
                                End If
                            Next
                            
                            For lngStartLine = lngStartLine1 To lngLastLine1
                                If vbc.CodeModule.Find("Public Function", lngStartLine, 1, lngStartLine, 15) = True Then
                                    lngFunction = lngFunction + 1
                                End If
                            Next
                            
                            For lngStartLine = lngStartLine1 To lngLastLine1
                                If vbc.CodeModule.Find("Private Function", lngStartLine, 1, lngStartLine, 16) = True Then
                                    lngFunction = lngFunction + 1
                                End If
                            Next

'Recorremos las líneas de código para localizar los comentarios
                            For lngStartLine = lngStartLine1 To lngLastLine1
                                If vbc.CodeModule.Find("'", lngStartLine, 1, lngStartLine, 10) = True Then lngComment = lngComment + 1
                            Next
                        End With
                        Exit For
                    Next
                End If
            End If
        Next
        
    Set vbcProjects = Nothing
    
    VBIDE_Estadisticas = Array(RefNum, NumMods, NumMod1, NumMod2, NumMod3, NumMod4, NumMod5, NumLinProcsCabecera, NumLinProcs, NumProcs, lngComment, lngSub, lngFunction)
       
End Function
