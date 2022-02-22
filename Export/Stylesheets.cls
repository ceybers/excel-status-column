VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stylesheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Stylesheets"
Option Explicit

Private Type TState
    Item As Scripting.Dictionary
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

'@DefaultMember
Public Property Get Item() As Scripting.Dictionary
Attribute Item.VB_UserMemId = 0
    Set Item = This.Item
End Property

Public Sub Load(ByVal filename As String)
    Dim stylesheetString As String
    Dim inputLine As String
    
    Open filename For Input As #1
    
    Do While Not EOF(1)
        Line Input #1, inputLine
        If Len(inputLine) > 0 Then
            If Left(inputLine, 1) <> " " Then
                AddStylesheet stylesheetString
                stylesheetString = inputLine & vbCrLf
            Else
                stylesheetString = stylesheetString & Trim(inputLine) & vbCrLf
            End If
        End If
    Loop
    
    AddStylesheet stylesheetString
    
    Close #1
End Sub

Public Sub AddStylesheet(ByVal stylesheetString As String)
    If Len(stylesheetString) = 0 Then Exit Sub
    
    If Len(stylesheetString) > 2 Then
        stylesheetString = Left(stylesheetString, Len(stylesheetString) - 2)
    End If
    
    Dim stylesheet As stylesheet
    Set stylesheet = New stylesheet
    
    Dim lines As Variant
    lines = Split(stylesheetString, vbCrLf)
    
    stylesheet.Name = lines(0)
        
    Dim i As Long
    For i = (LBound(lines) + 1) To UBound(lines)
        stylesheet.Add lines(i)
    Next i

    This.Item.Add Item:=stylesheet, key:=stylesheet.Name
End Sub
