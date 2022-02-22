VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stylesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Stylesheets"
Option Explicit

Private Type TState
    Name As String
    Item As Scripting.Dictionary
    Default As String
End Type
Private This As TState

Public Property Get Default() As String
    Default = This.Default
End Property

Public Property Let Default(ByVal Default As String)
    This.Default = Default
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal Name As String)
    This.Name = Name
End Property

'@DefaultMember
Public Property Get Item() As Scripting.Dictionary
Attribute Item.VB_UserMemId = 0
    Set Item = This.Item
End Property

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

Public Sub Add(ByVal inputLine As String)
    Dim splitQuotes As Variant
    splitQuotes = Split(inputLine, """")
        
    Dim ruleName As String
    Dim splitSpace As Variant
    If UBound(splitQuotes) = 0 Then
        splitSpace = Split(" " & inputLine, " ")
        ruleName = splitSpace(1)
    Else
        splitSpace = Split(" " & splitQuotes(2), " ")
        ruleName = splitQuotes(1)
    End If
    
    Dim rule As StylesheetRule
    Set rule = StylesheetRule.Create(ruleName, splitSpace(2), splitSpace(3))
    
    Dim i As Long
    For i = 4 To UBound(splitSpace)
        Select Case splitSpace(i)
            Case "Default"
                This.Default = ruleName
            Case "Bold"
                rule.bold = True
            Case "Italic"
                rule.italic = True
        End Select
    Next i

    This.Item.Add key:=rule.Name, Item:=rule
End Sub

Private Function Formula1() As String
    Dim result As String
    Dim rule As StylesheetRule
    Dim key As Variant
    
    For Each key In This.Item.Keys
        result = result & This.Item(key).Name & ","
    Next key
    
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If
    
    Formula1 = result
End Function

Public Sub AddFormatConditions(ByVal rng As Range)
    Dim rule As StylesheetRule
    Dim akey As Variant
    
    rng.FormatConditions.Delete
    
    For Each akey In This.Item.Keys
        AddFormatCondition rng, This.Item(akey)
    Next akey
End Sub

Private Sub AddFormatCondition(ByVal rng As Range, ByVal rule As StylesheetRule)
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & rule.Name & """"
    
    rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        .Color = rule.FontColor
        .bold = rule.bold
        .italic = rule.italic
        .TintAndShade = 0
    End With
    
    With rng.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = rule.BackgroundColor
        .TintAndShade = 0
    End With
End Sub

Public Sub AddValidation(ByVal rng As Range)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Formula1
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Public Sub ApplyDefault(ByVal rng As Range)
    If This.Default <> vbNullString Then
        rng.Value2 = This.Default
    End If
End Sub

Public Sub TryPrint(ByVal rng As Range)
    Dim c As Long
    c = This.Item.Count
    If c > rng.Rows.Count Then
        c = rng.Rows.Count
    End If
    
    Dim i As Long
    Dim key As Variant
    For Each key In This.Item.Keys
        i = i + 1
        If i <= c Then
            rng.Cells(i, 1).Value2 = This.Item(key).Name
        End If
    Next key
End Sub

Public Sub ApplyToRange(ByVal rng As Range, Optional DoApplyDefault As Boolean = True)
    AddValidation rng
    AddFormatConditions rng
    
    If DoApplyDefault Then ApplyDefault rng
    
    TryPrint rng
End Sub
