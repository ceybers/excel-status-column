VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleApplicatorViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SimpleApplicator")
Option Explicit
Implements IViewModel

Private Type TState
    Stylesheets As Stylesheets
    SelectedStylesheet As stylesheet
End Type
Private This As TState

Public Property Get SelectedStylesheet() As stylesheet
    Set SelectedStylesheet = This.SelectedStylesheet
End Property

Private Sub Class_Initialize()
    Set This.Stylesheets = New Stylesheets
    ' TODO Check if file exists, etc. & error handle this
    This.Stylesheets.Load STYLESHEET_FILENAME
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = False
End Function

Public Sub LoadToListBox(ByVal lb As MSForms.ListBox)
    lb.Clear
    
    Dim key As Variant
    For Each key In This.Stylesheets.Item.Keys
        lb.AddItem This.Stylesheets.Item(key).Name
    Next key
End Sub

Public Sub TrySelect(ByVal stylesheetName As String)
    Set This.SelectedStylesheet = This.Stylesheets.Item(stylesheetName)
End Sub
