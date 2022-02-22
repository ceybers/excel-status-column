VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdvancedApplicatorView 
   Caption         =   "Advanced Stylesheet Applicator"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   OleObjectBlob   =   "AdvancedApplicatorView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdvancedApplicatorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "AdvancedApplicator"
Option Explicit
Implements IView

Private vm As AdvancedApplicatorViewModel

Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub cmbOK_Click()
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As IViewModel) As Boolean
    Set vm = ViewModel
    This.IsCancelled = False
    
    vm.LoadToTreeView Me.tvStylesheets
    vm.LoadToListView Me.lvRules
    
    TrySelectFirstItems
    
    vm.LoadToDetails Me.frmRule
    vm.LoadToPreview Me.txtPreview
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub TrySelectFirstItems()
    'If Me.tvStylesheets.Nodes.Count > 0 Then
    '    Me.tvStylesheets.Nodes.Item(2).selected = True
    '    vm.TrySelectStylesheet Me.tvStylesheets.Nodes.Item(2).key
    '    vm.LoadToListView Me.lvRules
    'End If
    
    'If Me.lvRules.ListItems.Count > 0 Then
    '    Me.lvRules.ListItems.Item(1).selected = True
    '    vm.TrySelectStylesheetRule Me.lvRules.ListItems.Item(1).key
    'End If
End Sub

Private Sub tvStylesheets_NodeClick(ByVal Node As MSComctlLib.Node)
    vm.TrySelectStylesheet Node.key
    vm.LoadToListView Me.lvRules
    vm.LoadToDetails Me.frmRule
    vm.LoadToPreview Me.txtPreview
End Sub

Private Sub lvRules_ItemClick(ByVal Item As MSComctlLib.ListItem)
    vm.TrySelectStylesheetRule Item.key
    vm.LoadToDetails Me.frmRule
    vm.LoadToPreview Me.txtPreview
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

