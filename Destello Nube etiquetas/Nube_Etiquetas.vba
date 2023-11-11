
'Módulo estándar: modNubeEtiquetas'
Public Function DoCloud()
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-nube-de-etiquetas/
'                     Destello formativo 377
'--------------------------------------------------------------------------------------------------------
' Título            : DoCloud
' Autor original    : Philben
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : 28/11/2011
' Adaptado          : 18/05/2018
' Propósito         : crear la nube de etiquetas
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                     portapapeles y pega en el editor de VBA.
'
'Sub DoCloud_Test()
'
'   Call DoCloud
'
'End Sub
'-------------------------------------------------------------------------------------------------------
Dim rstTable As DAO.Recordset
Dim oTG
    
    On Error GoTo LinErr
    
    Set oTG = New clsTagCloud
    Set oTG.CtlSubForm = Form_EtiquetasContainer.EtiquetasCloud
    
        Set rstTable = CurrentDb.OpenRecordset("etiquetas", dbOpenSnapshot)
            Do Until rstTable.EOF
                oTG.AddTag rstTable!etiqnom, rstTable!etiqfrq, "Número de etiquetas : " & rstTable!etiqfrq
                rstTable.MoveNext
            Loop
        Set rstTable = Nothing
        
        With oTG
            .FontName = "Century Gothic"
            .setFontHexColors "#17365D", "#D6DFEC"
            .setFontWeights 700, 200
            .setMaxFontSize 24
            .setVerticalAlign eVerticalAlign.Baseline
            .setTagOrder 0 'Aleatorio
            .setOnHoverAttributes False, True, True
            .Go
        End With
     
    Form_EtiquetasContainer!EtiquetasCloud.visible = True
    
    Exit Function
    
LinErr:
    Form_EtiquetasContainer!EtiquetasCloud.visible = False

End Function

'Módulo de clase 1: clsTagCloud'
Option Compare Binary   'pour Like
Option Explicit

'---------------------------------------------------------------------------------------------
' Clase             : clsTagCloud v0.91
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-nube-de-etiquetas/
'                     Destello formativo 376
'--------------------------------------------------------------------------------------------------------
' Título            : DoCloud
' Autor original    : Philben
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : 28/11/2011
' Adaptado          : 18/05/2018
' Propósito         : gestionar la nube de etiquetas
'-----------------------------------------------------------------------------------------------------------------------------------------------

Public Event onClick(ByVal sMot As String)

Private Const gcsCtlTagName As String = "Tag"
Private Const gclMaxLong As Long = 2 ^ 31 - 1

Private Const gciMinFontWeight As Integer = 100
Private Const gciDefaultFontWeight As Integer = 400
Private Const gciMaxFontWeight As Integer = 900

Private Const gciMinFontSize As Long = 10
Private Const gciMaxFontSize As Long = 22

Private Const gclMinFontColor As Long = 0           'Black
Private Const gclDefaultFontColor As Long = 0
Private Const gclMaxFontColor As Long = &HFFFFFF    'White

Private Const gcsDefaultFontName As String = "Century Gothic"

Public Enum eVerticalAlign
   Baseline = 1
   Top
   Center
   Bottom
End Enum

Public Enum eOrderBy
   Aleatoire
  ' TagCroissant
  ' TagDecroissant
  ' FrequenceCroissant
  ' FrequenceDecroissant
End Enum

Private Type tRGB
   color As Long
   RED As Byte
   GREEN As Byte
   BLUE As Byte
End Type

Private Type tGeneralParameters
   sFontName As String
   iFontWeightFrom As Integer
   iFontWeightTo As Integer
   tFontColorFrom As tRGB
   tFontColorTo As tRGB
   eVertAlign As eVerticalAlign
   lNbTags As Long
   lMinFreq As Long
   lMaxFreq As Long
   iMaxFontSize As Integer
   eOrder As eOrderBy
   lBackColor As Long
   IsOnHoverUnderline As Boolean
   IsOnHoverBorder As Boolean
   IsOnHoverSwapColors As Boolean
   lInitFormWidth As Long
   lInitFormHeight As Long
End Type

Private WithEvents goSubFormSection As Access.Section
Private WithEvents goFormSection As Access.Section
Private gCollLabel As Collection
Private goCtlSubForm As Access.SubForm
Private glIdHover As Long, galIndex() As Long
Private gbIsActivate As Boolean
Private gtGenParams As tGeneralParameters

Private Sub Class_Initialize()
   
    Set gCollLabel = New Collection
       With gtGenParams
          .sFontName = gcsDefaultFontName
          .tFontColorFrom = LongToRGB(gclDefaultFontColor)
          .tFontColorTo = LongToRGB(gclDefaultFontColor)
          .iFontWeightFrom = gciDefaultFontWeight
          .iFontWeightTo = gciDefaultFontWeight
          .lMinFreq = gclMaxLong
          .iMaxFontSize = gciMaxFontSize
          .eOrder = eOrderBy.Aleatoire
          .eVertAlign = Bottom
          .lBackColor = HexToRGB("#FAFBFC").color
       End With

End Sub

Private Sub Class_Terminate()

    gbIsActivate = False
    
    Set goSubFormSection = Nothing
    
    Set goFormSection = Nothing
    
    Set goCtlSubForm = Nothing
    
    Set gCollLabel = Nothing

End Sub

'Establecemos la fuente de las etiquetas
Public Property Let FontName(ByVal sFontName As String)
Dim I As Long
   
    gtGenParams.sFontName = sFontName
       
    For I = 1 To gCollLabel.Count
        gCollLabel(I).goLabel.FontName = sFontName
    Next I

End Property

Public Property Set CtlSubForm(oCtlSubForm As Access.SubForm)

    Set goCtlSubForm = oCtlSubForm
   
    With gtGenParams
       .lInitFormHeight = goCtlSubForm.Height
       .lInitFormWidth = goCtlSubForm.Width
    End With
   
End Property

Public Property Get IsActivate()
   
    IsActivate = gbIsActivate

End Property

Public Function AddTag(ByVal sText As String, ByVal lFrequency As Long, ByVal sTipText As String) As Boolean
   
    If Not gbIsActivate Then
        If gCollLabel.Count = 0 Then
             With goCtlSubForm
                .Form.InsideWidth = .Width
                .Form.InsideHeight = .Height
                .Form.visible = False
             End With
             CountCtlTags
        End If
        
        With gtGenParams
            If .lNbTags < gCollLabel.Count Then
                .lNbTags = .lNbTags + 1
                If lFrequency < .lMinFreq Then .lMinFreq = lFrequency
                If lFrequency > .lMaxFreq Then .lMaxFreq = lFrequency
                gCollLabel(.lNbTags).SetInfos sText, lFrequency, sTipText
                AddTag = True
            End If
        End With
    End If

End Function

Public Sub setBackColor(ByVal sColor As String)
   
    gtGenParams.lBackColor = HexToRGB(sColor).color

End Sub

Public Property Get BackColor() As Long
   
    BackColor = gtGenParams.lBackColor

End Property

Public Sub setOnHoverAttributes(ByVal Underline As Boolean, ByVal Border As Boolean, ByVal SwapColors As Boolean)
   
    With gtGenParams
       .IsOnHoverUnderline = Underline
       .IsOnHoverBorder = Border
       .IsOnHoverSwapColors = SwapColors
    End With
   
End Sub

Public Sub setMaxFontSize(ByVal iSize As Integer)
   
    If iSize >= gciMinFontSize And iSize <= gciMaxFontSize Then gtGenParams.iMaxFontSize = iSize

End Sub

Public Sub setFontLongColors(ByVal lFrom As Long, ByVal lTo As Long)
   
    With gtGenParams
       .tFontColorFrom = LongToRGB(lFrom)
       .tFontColorTo = LongToRGB(lTo)
    End With
   
End Sub

Public Sub setFontHexColors(ByVal sFrom As String, ByVal sTo As String)
   
    With gtGenParams
       .tFontColorFrom = HexToRGB(sFrom)
       .tFontColorTo = HexToRGB(sTo)
    End With
   
End Sub

Public Sub setFontWeights(ByVal iFrom As Integer, ByVal iTo As Integer)
   
    With gtGenParams
       If iFrom >= gciMinFontWeight And iFrom <= gciMaxFontWeight Then .iFontWeightFrom = iFrom
       If iTo >= gciMinFontWeight And iTo <= gciMaxFontWeight Then .iFontWeightTo = iTo
    End With
   
End Sub

Public Sub setVerticalAlign(ByVal eType As eVerticalAlign)
   
    gtGenParams.eVertAlign = eType

End Sub

Public Sub setTagOrder(ByVal eOrder As eOrderBy)
   
    gtGenParams.eOrder = eOrder

End Sub

Public Property Get IsOnHoverUnderline() As Boolean
   
    IsOnHoverUnderline = gtGenParams.IsOnHoverUnderline

End Property

Public Property Get IsOnHoverBorder() As Boolean
   
    IsOnHoverBorder = gtGenParams.IsOnHoverBorder

End Property

Public Property Get IsOnHoverSwapColors() As Boolean
   
    IsOnHoverSwapColors = gtGenParams.IsOnHoverSwapColors

End Property

Public Sub Go()
Dim I As Long

    DoIndex
    
    ComputeColors
    ComputeWeights
    ComputeSizes
    ComputePositions
    
    For I = 1 To gtGenParams.lNbTags
        gCollLabel(I).Activate
    Next I
    
    Set oSubFormSection = goCtlSubForm.Form.Section(acDetail)
    Set oFormSection = goCtlSubForm.Parent.Section(acDetail)
    
    goCtlSubForm.Form.Section(acDetail).BackColor = gtGenParams.lBackColor
    goCtlSubForm.Form.visible = True
    gbIsActivate = True

End Sub

Private Property Set oSubFormSection(ByRef oSection As Access.Section)

    Set goSubFormSection = oSection
    
    goSubFormSection.OnMouseMove = "[Event Procedure]"
    
End Property

Private Sub CountCtlTags()
Dim oTCLabel As clsTagCloudLabel
Dim oCtl As Access.control
Dim lCount As Long, lLen As Long

    lLen = Len(gcsCtlTagName)
    
    For Each oCtl In goCtlSubForm.Form.Section(acDetail).Controls
        If oCtl.ControlType = acLabel And Left$(oCtl.Name, lLen) = gcsCtlTagName Then
            lCount = lCount + 1
            Set oTCLabel = New clsTagCloudLabel
                oTCLabel.Init lCount, Me, oCtl
                gCollLabel.Add oTCLabel
        End If
    Next oCtl
   
End Sub

Private Sub goSubFormSection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If glIdHover > 0 Then
        gCollLabel(glIdHover).NoHover
        glIdHover = 0
    End If

End Sub

Private Property Set oFormSection(ByRef oSection As Access.Section)

    Set goFormSection = oSection
    goFormSection.OnMouseMove = "[Event Procedure]"

End Property

Private Sub goFormSection_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    If glIdHover > 0 Then
        gCollLabel(glIdHover).NoHover
        glIdHover = 0
    End If
    
    err.Clear

End Sub

Friend Sub onHover(ByVal lId As Long)

    If glIdHover > 0 Then
        gCollLabel(glIdHover).NoHover
    End If
       
    glIdHover = lId

End Sub

Friend Sub onClick(ByVal sTag As String)
Dim sFiltro As String

    sFiltro = "correoeti01='" & sTag & "'"
    sFiltro = sFiltro & " OR " & "correoeti02='" & sTag & "'"
    sFiltro = sFiltro & " OR " & "correoeti03='" & sTag & "'"
    sFiltro = sFiltro & " OR " & "correoeti04='" & sTag & "'"
    sFiltro = sFiltro & " OR " & "correoeti05='" & sTag & "'"
    sFiltro = sFiltro & " OR " & "correoeti06='" & sTag & "'"
    
    If Form_EtiquetasContainer.visible = True Then
        Form_EtiquetasContainer.Filter = sFiltro
        Form_EtiquetasContainer.FilterOn = True
    Else
        Form_EtiquetasContainer.Filter = sFiltro
        Form_EtiquetasContainer.FilterOn = True
    End If
    
    'Call DoCloud
    
    'RaiseEvent onClick(lId, sTag)

End Sub

Private Sub ComputeColors()
Dim I As Long
Dim dCstRatio As Double, dRatio As Double
Dim lColor As Long
Dim bCalcRatio As Boolean

    With gtGenParams
        If .lMaxFreq > .lMinFreq And .tFontColorFrom.color <> .tFontColorTo.color Then
            dCstRatio = CDbl(.lMaxFreq - .lMinFreq)
            bCalcRatio = True
        Else
            lColor = .tFontColorFrom.color
        End If
    End With
    
    For I = 1 To gtGenParams.lNbTags
        With gCollLabel(I)
            If bCalcRatio Then
                dRatio = (.Frequency - gtGenParams.lMinFreq) / dCstRatio
                .setForeColor = getGradientColor(dRatio)
            Else
                .setForeColor = lColor
            End If
        End With
    Next I

End Sub

Private Sub ComputeWeights()
Dim I As Long
Dim dCstRatio As Double
Dim iWeight As Integer
Dim bCalcRatio As Boolean

    With gtGenParams
        If .iFontWeightFrom < .iFontWeightTo Then
            iWeight = .iFontWeightFrom
        Else
            iWeight = .iFontWeightTo
        End If
        If .lMaxFreq > .lMinFreq And .iFontWeightFrom <> .iFontWeightTo Then
            dCstRatio = CDbl(Abs(.iFontWeightFrom - .iFontWeightTo)) / (.lMaxFreq - .lMinFreq)
            bCalcRatio = True
        End If
    End With
    
    For I = 1 To gtGenParams.lNbTags
        With gCollLabel(I)
            If bCalcRatio Then
                .goLabel.FontWeight = Int(iWeight + (.Frequency - gtGenParams.lMinFreq) * dCstRatio)
            Else
                .goLabel.FontWeight = iWeight
            End If
        End With
    Next I

End Sub

Private Sub ComputeSizes()
Const clMinSpace As Long = 100
Dim oTCLabel As clsTagCloudLabel
Dim dRatio As Double
Dim lFormWidth As Long, I As Long, lId As Long
Dim iMinFontSize As Integer, iMaxFontSize As Integer
Dim bCalcRatio As Boolean

   With gtGenParams
      iMaxFontSize = .iMaxFontSize
      iMinFontSize = gciMinFontSize
      lFormWidth = .lInitFormWidth - 50
      If .lMaxFreq > .lMinFreq Then
         bCalcRatio = True
         dRatio = CDbl(iMaxFontSize - iMinFontSize) / (.lMaxFreq - .lMinFreq)
      Else
         dRatio = 1
      End If
   End With

   For I = 1 To gtGenParams.lNbTags
      Set oTCLabel = gCollLabel(galIndex(I))
      With oTCLabel
         .CalculSize iMinFontSize, gtGenParams.lMinFreq, dRatio, clMinSpace
         If .Width + clMinSpace > lFormWidth Then
            With gtGenParams
               If iMaxFontSize > gciMinFontSize Then
                  iMaxFontSize = Int(iMaxFontSize * (CDbl(lFormWidth / (oTCLabel.Width + clMinSpace))))
                  If iMaxFontSize < gciMinFontSize Then iMaxFontSize = gciMinFontSize
                  If bCalcRatio Then
                     dRatio = CDbl(iMaxFontSize - iMinFontSize) / (.lMaxFreq - .lMinFreq)
                  Else
                     iMinFontSize = iMaxFontSize
                  End If
                  lId = I
               End If
            End With
         End If
      End With
   Next I

   For I = 1 To lId
      gCollLabel(galIndex(I)).CalculSize iMinFontSize, gtGenParams.lMinFreq, dRatio, clMinSpace
   Next I

End Sub

Private Sub ComputePositions()
Const clMargeX As Long = 25
Const clMargeY As Long = 20
Dim I As Long, X As Long, Y As Long, lLblPerLine As Long
Dim lLineHeight As Long, lFormWidth As Long

    goCtlSubForm.Form.InsideWidth = gtGenParams.lInitFormWidth
    lFormWidth = goCtlSubForm.Form.InsideWidth
    
    lLblPerLine = 0
    X = clMargeX
    Y = clMargeY

    For I = 1 To gtGenParams.lNbTags
       With gCollLabel(galIndex(I))
          If X + .Width + 1 > lFormWidth Then
             UpdateHeight Y + lLineHeight + clMargeY
             setLineTags I, lFormWidth, Y, lLineHeight, lLblPerLine, X, clMargeX
             X = clMargeX
             lLblPerLine = 0
             Y = Y + lLineHeight + clMargeY
             lLineHeight = .Height
          ElseIf .Height > lLineHeight Then
             lLineHeight = .Height
          End If
          X = X + .Width + clMargeX
          lLblPerLine = lLblPerLine + 1
       End With
    Next I
   
    UpdateHeight Y + lLineHeight + clMargeY
    setLineTags I, lFormWidth, Y, lLineHeight, lLblPerLine, X, clMargeX

End Sub

Private Sub UpdateHeight(ByVal lHeight As Long)
   
    If goCtlSubForm.Parent.Section(acDetail).Height < goCtlSubForm.Top + lHeight Then
        goCtlSubForm.Parent.Section(acDetail).Height = goCtlSubForm.Top + lHeight
    End If
       
    goCtlSubForm.Height = lHeight
       
    goCtlSubForm.Form.Section(acDetail).Height = goCtlSubForm.Height

End Sub

Private Sub setLineTags(ByVal lIdCurTag As Long, ByVal lFormWidth As Long, ByVal Y As Long, _
                        ByVal lLineHeight As Long, ByVal lNbTags As Long, ByVal xMax As Long, ByVal lMargeX As Long)
Dim xSpace As Long, X As Long, I As Long
    
    xSpace = CDbl(lFormWidth - xMax) / lNbTags
    X = lMargeX
    
    For I = lIdCurTag - lNbTags To lIdCurTag - 1
        With gCollLabel(galIndex(I))
            .SetPos X, Y, xSpace, lLineHeight, gtGenParams.eVertAlign
            X = X + .Width + lMargeX
        End With
    Next I
    
End Sub

Private Function LongToRGB(ByVal lColor As Long) As tRGB
   
    If lColor >= gclMinFontColor And lColor <= gclMaxFontColor Then
        With LongToRGB
            .RED = lColor Mod &H100
            .GREEN = (lColor \ &H100) Mod &H100
            .BLUE = (lColor \ &H10000) Mod &H100
            .color = lColor
        End With
    End If

End Function

Private Function getGradientColor(ByVal dRatio As Double) As Long
Dim tColor As tRGB

    With gtGenParams
        tColor.RED = .tFontColorFrom.RED * dRatio + .tFontColorTo.RED * (1 - dRatio)
        tColor.GREEN = .tFontColorFrom.GREEN * dRatio + .tFontColorTo.GREEN * (1 - dRatio)
        tColor.BLUE = .tFontColorFrom.BLUE * dRatio + .tFontColorTo.BLUE * (1 - dRatio)
    End With
    
    getGradientColor = rgb(tColor.RED, tColor.GREEN, tColor.BLUE)

End Function

Private Function HexToRGB(ByVal sHexColor As String) As tRGB

    sHexColor = Replace(Trim$(sHexColor), "#", "")
       If Len(sHexColor) = 6 Then
          With HexToRGB
             .RED = val("&H" & Left$(sHexColor, 2))
             .GREEN = val("&H" & Mid$(sHexColor, 3, 2))
             .BLUE = val("&H" & Right$(sHexColor, 2))
             .color = rgb(.RED, .GREEN, .BLUE)
          End With
       End If

End Function

Private Sub DoIndex()
Dim I As Long, J As Long, K As Long, lMax As Long

    lMax = gtGenParams.lNbTags
    ReDim galIndex(1 To lMax)
    
    For I = 1 To lMax
        galIndex(I) = I
    Next I
    
    If gtGenParams.eOrder = eOrderBy.Aleatoire Then
        Randomize
        For I = 1 To lMax - 1
            J = Int((lMax - I + 1) * Rnd) + I
            K = galIndex(I)
            galIndex(I) = galIndex(J)
            galIndex(J) = K
        Next I
    Else
        ShellSortIndex 1, lMax
    End If

End Sub

'Devuelve el ínidice de las etiquetas
Public Sub ShellSortIndex(ByVal lLowerBound As Long, ByVal lUpperBound As Long)
Dim I As Long, lIdx As Long, lRefIdx As Long, lInc As Long, N As Long, lMin As Long, lGapIdx As Long
Dim vRefVal As Variant, avGap As Variant

    N = lUpperBound - lLowerBound + 1
    avGap = VBA.Array(1, 4, 10, 23, 57, 132, 301, 701, 1750, 4254, 10321, 25040, 60748, 147376, 357535, 867381, 2104267, 5104953)
    lGapIdx = UBound(avGap)
    
    While avGap(lGapIdx) >= N: lGapIdx = lGapIdx - 1: Wend

    With gCollLabel
       While lGapIdx >= 0
          lInc = avGap(lGapIdx)
          lMin = lLowerBound + lInc
          For I = lMin To lUpperBound
             lIdx = I
             lRefIdx = galIndex(lIdx)
                vRefVal = .Item(lRefIdx).Frequency
             galIndex(lIdx) = lRefIdx
          Next I
          lGapIdx = lGapIdx - 1
       Wend
    End With
   
End Sub

'Módulo de clase 2: clsTagCloudlabel'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------
' Clase             : clsTagCloudLabel v0.9
'--------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-nube-de-etiquetas/
'                     Destello formativo 376
'--------------------------------------------------------------------------------------------------------
' Título            : DoCloud
' Autor original    : Philben
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : 28/11/2011
' Adaptado          : 18/05/2018
' Propósito         : gestionar las etiquetas de la nube
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Const gcsActiveEvent As String = "[Event Procedure]"

Public WithEvents goLabel As Access.Label
Private goTagCloud As clsTagCloud

Private gbOver As Boolean
Private glId As Long, glFrequency As Long, glFontSize As Long, glCtlWidth As Long, glCtlHeight As Long, glForeColor As Long

Private Sub Class_Terminate()
   
    With goLabel
       .visible = False
       .OnMouseMove = vbNullString
    End With

End Sub

Public Property Get Width() As Long
   
    Width = glCtlWidth

End Property

Public Property Get Height() As Long
   
    Height = glCtlHeight

End Property

Public Property Get Frequency() As Long
   
    Frequency = glFrequency

End Property

Public Property Get Tag()
   
    Tag = goLabel.Caption

End Property

Public Sub Init(ByVal lId As Long, ByRef oParent As clsTagCloud, ByRef oLabel As Access.Label)

    glId = lId
    
    Set goTagCloud = oParent
    Set goLabel = oLabel
       
    With goLabel
        .OnMouseMove = vbNullString
        .visible = False
        .FontUnderline = False
        .FontItalic = False
        .BorderStyle = 0
        .BackStyle = 0
    End With
       
End Sub

Public Sub SetInfos(ByVal sText As String, ByVal lFrequency As Long, ByVal sTipText As String)
   
    goLabel.Caption = sText
    goLabel.ControlTipText = sTipText
    glFrequency = lFrequency

End Sub

Public Property Let setForeColor(ByVal lForeColor As Long)
   
    glForeColor = lForeColor

End Property

Public Sub NoHover()
   
    With goLabel
       .FontUnderline = False
       If goTagCloud.IsOnHoverSwapColors Then
          .ForeColor = glForeColor
          .BackColor = goTagCloud.BackColor
          .BackStyle = 0
       End If
       .BorderStyle = 0
    End With
    
    gbOver = False
   
End Sub

Public Sub CalculSize(ByVal lMinFontSize As Long, ByVal lMinFreq As Long, ByVal dRatio As Double, ByVal lMinSpace As Long)
Const clMinHeight As Long = 40
Const clKey As Long = 51488399

    glFontSize = Int(lMinFontSize + (glFrequency - lMinFreq) * dRatio)
        
    WizHook.Key = clKey
    
    With goLabel
        Call WizHook.TwipsFromFont(.FontName, glFontSize, .FontWeight, False, _
                                   goTagCloud.IsOnHoverUnderline, 0, .Caption, 0, glCtlWidth, glCtlHeight)
    End With
    
    glCtlWidth = glCtlWidth + lMinSpace
       
    glCtlHeight = glCtlHeight + clMinHeight
       
End Sub

Public Sub SetPos(ByVal X As Long, ByVal Y As Long, ByVal xSpace As Long, ByVal Height As Long, ByVal eVertAlign As eVerticalAlign)

    glCtlWidth = glCtlWidth + xSpace

    With goLabel
       .Left = X
       .Width = glCtlWidth
       .Height = glCtlHeight

        Select Case eVertAlign
            Case eVerticalAlign.Baseline
                .Top = Y + (Height - glCtlHeight) / 1.23
            Case eVerticalAlign.Bottom
                .Top = Y + Height - glCtlHeight
            Case eVerticalAlign.Center
                .Top = Y + (Height - glCtlHeight) / 2
            Case eVerticalAlign.Top
                .Top = Y
        End Select
        
    End With
   
End Sub

Public Sub Activate()
   
    With goLabel
      .FontSize = glFontSize
      .ForeColor = glForeColor
      .visible = True
      .OnMouseMove = gcsActiveEvent
   End With
   
End Sub

Private Sub goLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If goTagCloud.IsActivate Then
        With goLabel
            If Not gbOver Then
                gbOver = True
                .FontUnderline = goTagCloud.IsOnHoverUnderline
                If goTagCloud.IsOnHoverBorder Then
                    .BorderStyle = 1
                    .BorderWidth = 1
                    .BorderColor = 0
                End If
                
                If goTagCloud.IsOnHoverSwapColors Then
                    .BackColor = glForeColor
                    .ForeColor = goTagCloud.BackColor
                    .BackStyle = 1
                End If
                goTagCloud.onHover glId
            End If
    
            If Button = acLeftButton Then
                Application.Echo False
                DoCmd.Hourglass True
                   goTagCloud.onClick goLabel.Caption
                DoCmd.Hourglass False
                Application.Echo True
            End If
        End With
    End If

End Sub