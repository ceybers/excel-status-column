Attribute VB_Name = "TestSimpleApplicator"
'@Folder("SimpleApplicator")
Option Explicit

Public Sub TestSimpleApplicator()
    Dim vm As SimpleApplicatorViewModel
    Dim view As IView
    
    Set vm = New SimpleApplicatorViewModel
    
    Set view = New SimpleApplicatorView
    
    If view.ShowDialog(vm) Then
        Dim lc As ListColumn
        For Each lc In GetListColumnsFromSelection(Selection)
            vm.SelectedStylesheet.ApplyToRange lc.DataBodyRange
        Next lc
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub
