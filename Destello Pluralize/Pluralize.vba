Option Compare Database
Option Explicit


'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-funcion-pluralize-by-mike-wolfe/
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Pluralize
' Autor original    : Mike Wolfe <mike@nolongerset.com>  (10/21/2010 - 7/24/2014)
' Adaptado por      : Luis Viadel | https://cowtechnologies.net
' Actualizado       : 03/04/2022 (Luis Viadel)  
' Propósito         : Formats a phrase to make verbs agree in number.
'-----------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : Pluralize
' Author    : Mike Wolfe <mike@nolongerset.com>
' Date      : 10/21/2010 - 7/24/2014
' Adapted   : Luis Viadel | https://cowtechnologies.net
' Date      : 03/04/2022
' Purpose   : Formats a phrase to make verbs agree in number.
' Notes     : To substitute the absolute value of the number for numbers that can be
'               positive or negative, use a custom number format that includes
'               both positive and negative formats; e.g., "#;#".
' Usage     : Msg = "There [is/are] # record[s].  [It/They] consist[s/] of # part[y/ies] each."
'>>> Pluralize("There [is/are] # record[s].  [It/They] consist[s/] of # part[y/ies] each.", 1)
' There is 1 record.  It consists of 1 party each.
'>>> Pluralize("There [is/are] # record[s].  [It/They] consist[s/] of # part[y/ies] each.", 6)
' There are 6 records.  They consist of 6 parties each.
'>>> Pluralize("There was a {gain/loss} of # dollar[s].", -50, "#", "#;#")
' There was a loss of 50 dollars.
'>>> Pluralize("I {won/lost} # at the fair.  {I was thrilled./I'll never learn.}", 20, "#", "Currency")
' I won $20.00 at the fair.  I was thrilled.
'>>> Pluralize("There [is/are] # {more/less} finger[s] on his hand after the surgery.", -1, "#", "#;#")
' There is 1 less finger on his hand after the surgery.
'---------------------------------------------------------------------------------------
' Adaptation purpose  : adapt the function to the Spanish language
'                        in English: Pluralize("You gambled your life savings and {won/won/lost} #.",i, ,"#,##0 €;#,##0 €;nothing"):Next i
'                        in spanish: Pluralize("Apostaste los ahorros de tu vida y {ganaste/no ganaste/perdiste} #.",i, ,"#,##0 €;#,##0 €;nada"):Next i
'---------------------------------------------------------------------------------------

Function Pluralize(Text As String, Num As Variant, _
                   Optional NumToken As String = "#", _
                   Optional NumFormat As String = "")
    
    Const OpeningBracket As String = "\["
    Const ClosingBracket As String = "\]"
    Const OpeningBrace As String = "\{"
    Const ClosingBrace As String = "\}"
    Const DividingSlash As String = "/"
    Const CharGroup As String = "([^\]]*)"  'Group of 0 or more characters not equal to closing bracket
    Const BraceGroup As String = "([^\/\}]*)" 'Group of 0 or more characters not equal to closing brace or dividing slash

    Dim IsPlural As Boolean, IsNegative As Boolean, IsZero As Boolean
    
    If IsNumeric(Num) Then
        IsPlural = (Abs(Num) <> 1)
        IsNegative = (Num < 0)
        IsZero = (Num = 0)
    End If
    
    Dim Msg As String, Pattern As String
    Msg = Text
    
    'Replace the number token with the actual number
    Msg = Replace(Msg, NumToken, Format(Num, NumFormat))
    
    'Replace [y/ies] style references
    Pattern = OpeningBracket & CharGroup & DividingSlash & CharGroup & ClosingBracket
    Msg = RegExReplace(Pattern, Msg, "$" & IIf(IsPlural, 2, 1))
    
    'Replace [s] style references
    Pattern = OpeningBracket & CharGroup & ClosingBracket
    Msg = RegExReplace(Pattern, Msg, IIf(IsPlural, "$1", ""))
        
    'Replace {gain/loss} style references
    Pattern = OpeningBrace & BraceGroup & DividingSlash & BraceGroup & DividingSlash & BraceGroup & ClosingBrace
    Msg = RegExReplace(Pattern, Msg, "$" & IIf(IsZero, 2, (IIf(IsNegative, 3, 1))))
       
    Pluralize = Msg
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RegExReplace
' Author    : Mike Wolfe <mike@nolongerset.com>
' Date      : 11/4/2010
' Source    : https://nolongerset.com/now-you-have-two-problems/
' Purpose   : Attempts to replace text in the TextToSearch with text and back references
'               from the ReplacePattern for any matches found using SearchPattern.
' Notes     - If no matches are found, TextToSearch is returned unaltered.  To get
'               specific info from a string, use RegExExtract instead.
'>>> RegExReplace("(.*)(\d{3})[\)\s.-](\d{3})[\s.-](\d{4})(.*)", "My phone # is 570.555.1234.", "$1($2)$3-$4$5")
'My phone # is (570)555-1234.
'---------------------------------------------------------------------------------------
'
Function RegExReplace(SearchPattern As String, TextToSearch As String, ReplacePattern As String, _
                      Optional GlobalReplace As Boolean = True, _
                      Optional IgnoreCase As Boolean = False, _
                      Optional MultiLine As Boolean = False) As String
Dim RE As Object

    Set RE = CreateObject("vbscript.regexp")
    With RE
        .MultiLine = MultiLine
        .Global = GlobalReplace
        .IgnoreCase = IgnoreCase
        .Pattern = SearchPattern
    End With
    
    RegExReplace = RE.Replace(TextToSearch, ReplacePattern)
    
End Function
