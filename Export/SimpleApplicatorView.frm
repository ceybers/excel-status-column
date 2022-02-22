VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SimpleApplicatorView 
   Caption         =   "Apply Structured Column"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3315
   OleObjectBlob   =   "SimpleApplicatorView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SimpleApplicatorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SimpleApplicator"
Option Explicit
Implements IView

Private vm As SimpleApplicatorViewModel

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
    
    vm.LoadToListBox Me.lbStylesheets
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub lbStylesheets_Change()
    vm.TrySelect Me.lbStylesheets
    Me.cmbOK.Enabled = True
End Sub

Private Sub lbStylesheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    vm.TrySelect Me.lbStylesheets
    Me.Hide
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
