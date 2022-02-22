Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Public Sub test()
    Dim Stylesheets As Stylesheets
    Set Stylesheets = New Stylesheets
    Stylesheets.Load "Stylesheet.txt"
    
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim lc As ListColumn
    Set lc = lo.ListColumns(3)
    
    Dim target As Range
    Set target = lc.DataBodyRange
    
    If RangeIsEmpty(target) Or RangeHasValidation(target) Or RangeHasFormatConditions(target) Then
        target.ClearFormats ' Clears manually-set (non-conditional) formatting only
        target.ClearContents
    End If

    With Stylesheets("Green Status")
        .AddValidation target
        .AddFormatConditions target
        .TryPrint target
        .ApplyDefault target
    End With
End Sub


