Attribute VB_Name = "TestAdvancedApplicator"
'@Folder("AdvancedApplicator")
Option Explicit

Public Sub TestAdvancedApplicator()
    Dim vm As AdvancedApplicatorViewModel
    Dim view As IView
    
    Set vm = New AdvancedApplicatorViewModel
    
    Set view = New AdvancedApplicatorView
    
    If view.ShowDialog(vm) Then
        Dim lc As ListColumn
        For Each lc In GetListColumnsFromSelection(Selection)
            vm.SelectedStylesheet.ApplyToRange lc.DataBodyRange
        Next lc
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub
