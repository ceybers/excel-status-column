VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdvancedApplicatorViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AdvancedApplicator")
Option Explicit
Implements IViewModel

Private Type TState
    Stylesheets As Stylesheets
    SelectedStylesheet As stylesheet
    SelectedStylesheetRule As StylesheetRule
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

' TODO Change to ListView
Public Sub LoadToTreeView(ByVal tv As MSComctlLib.TreeView)
    Dim parentNode As Node
    
    tv.Nodes.Clear
    tv.Appearance = cc3D
    tv.FullRowSelect = True
    tv.HideSelection = False
    tv.Indentation = 14
    tv.LineStyle = tvwRootLines
    tv.Style = tvwTreelinesPlusMinusText
    
    Set parentNode = tv.Nodes.Add(key:="Root", Text:="Stylesheets")
    parentNode.Expanded = True
    
    Dim key As Variant
    For Each key In This.Stylesheets.Item.Keys
        tv.Nodes.Add parentNode, tvwChild, key, key
    Next key
    
    If tv.Nodes.Count > 0 Then
        tv.Nodes.Item(2).selected = True
        Set This.SelectedStylesheet = This.Stylesheets(tv.Nodes.Item(2).key)
    End If
End Sub

Public Sub LoadToListView(ByVal lv As MSComctlLib.ListView)
    lv.view = lvwReport
    lv.Gridlines = True
    lv.Appearance = cc3D
    lv.FullRowSelect = True
    lv.HideSelection = False
    lv.LabelEdit = lvwManual
    
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    lv.ColumnHeaders.Add Text:="Rule", Width:=64
    lv.ColumnHeaders.Add Text:="Font", Width:=64
    lv.ColumnHeaders.Add Text:="Background", Width:=64
    
    If This.SelectedStylesheet Is Nothing Then
        Exit Sub
    End If
    
    Dim curNode As Node
    Dim key As Variant
    For Each key In This.SelectedStylesheet.Item.Keys
        'Set curNode = lv.ListItems.Add(key:=key, Text:=key)
        lv.ListItems.Add key:=key, Text:=key
        lv.ListItems(key).ListSubItems.Add Text:=This.SelectedStylesheet.Item(key).FontColor
        lv.ListItems(key).ListSubItems.Add Text:=This.SelectedStylesheet.Item(key).BackgroundColor
    Next key
    
    If This.SelectedStylesheet Is Nothing Then
        Exit Sub
    End If
    
    If lv.ListItems.Count > 0 Then
        lv.ListItems.Item(1).selected = True
        Set This.SelectedStylesheetRule = This.SelectedStylesheet.Item(lv.ListItems.Item(1).key)
    End If
End Sub

Public Sub LoadToListBox(ByVal lb As MSForms.ListBox)
    lb.Clear
    
    Dim key As Variant
    For Each key In This.Stylesheets.Item.Keys
        lb.AddItem This.Stylesheets.Item(key).Name
    Next key
End Sub

Public Sub LoadToDetails(ByVal frame As MSForms.frame)
    If This.SelectedStylesheetRule Is Nothing Then
        Exit Sub
    End If
    
    With This.SelectedStylesheetRule
        frame.Controls("txtRuleName") = .Name
        frame.Controls("txtFontColor") = .FontColor
        frame.Controls("txtBackgroundColor") = .BackgroundColor
        frame.Controls("chkBold") = .bold
        frame.Controls("chkItalic") = .italic
        frame.Controls("chkDefault") = (This.SelectedStylesheet.Default = .Name)
    End With
End Sub

Public Sub LoadToPreview(ByVal tb As MSForms.TextBox)
    If This.SelectedStylesheetRule Is Nothing Then
        Exit Sub
    End If
    
    With tb
        .Value = This.SelectedStylesheetRule.Name
        .ForeColor = This.SelectedStylesheetRule.FontColor
        .BackColor = This.SelectedStylesheetRule.BackgroundColor
        .Font.bold = This.SelectedStylesheetRule.bold
        .Font.italic = This.SelectedStylesheetRule.italic
    End With
End Sub

Public Sub TrySelectStylesheet(ByVal stylesheetName As String)
    Set This.SelectedStylesheet = This.Stylesheets.Item(stylesheetName)
End Sub

Public Sub TrySelectStylesheetRule(ByVal stylesheetNameRule As String)
    If This.SelectedStylesheet Is Nothing Then
        Exit Sub
    End If
    Set This.SelectedStylesheetRule = This.SelectedStylesheet.Item(stylesheetNameRule)
End Sub

