VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StylesheetRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Stylesheets"
Option Explicit

Private Type TState
    Name As String
    FontColor As Long
    BackgroundColor As Long
    bold As Boolean
    italic As Boolean
    Default As Boolean
End Type
Private This As TState

Public Property Get italic() As Boolean
    italic = This.italic
End Property

Public Property Let italic(ByVal RHS As Boolean)
    This.italic = RHS
End Property

Public Property Get bold() As Boolean
    bold = This.bold
End Property

Public Property Let bold(ByVal RHS As Boolean)
    This.bold = RHS
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get FontColor() As Long
    FontColor = This.FontColor
End Property

Public Property Let FontColor(ByVal RHS As Long)
    This.FontColor = RHS
End Property

Public Property Get BackgroundColor() As Long
    BackgroundColor = This.BackgroundColor
End Property

Public Property Let BackgroundColor(ByVal RHS As Long)
    This.BackgroundColor = RHS
End Property

Public Sub SetValues(ByVal Name As String, ByVal FontColor As Long, ByVal BackgroundColor As Long)
    With This
        .Name = Name
        .FontColor = FontColor
        .BackgroundColor = BackgroundColor
    End With
End Sub

Public Function Create(ByVal Name As String, ByVal FontColor As Long, ByVal BackgroundColor As Long) As StylesheetRule
    Dim result As StylesheetRule
    Set result = New StylesheetRule
    result.SetValues Name, FontColor, BackgroundColor
    Set Create = result
End Function
