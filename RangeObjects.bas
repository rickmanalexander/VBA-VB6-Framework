Attribute VB_Name = "RangeObjects"
'@Folder("Framework.CommonMethods")
Option Explicit
Option Private Module

Public Sub ClearCellContents(ByRef cellRange As Range)
    
    'is a merged cell, then clear this way
    'other wise you wil receive an error
    If cellRange.MergeCells Then
        cellRange.MergeArea.ClearContents
        
    Else
        cellRange.ClearContents

    End If
    
End Sub

Public Sub ClearRangeFilter(ByRef WrkSht As Worksheet, ByVal rangeBeginAddress As String, _
                            ByVal rangeEndAddress As String)
        
    With WrkSht
        If Not .AutoFilterMode Then Exit Sub
        
        'Clear it
        .Range(rangeBeginAddress & ":" & rangeEndAddress).AutoFilter

        'Add it back
        .Range(rangeBeginAddress & ":" & rangeEndAddress).AutoFilter
        
    End With
        
End Sub

Public Function TableExists(ByRef WrkShtIn As Worksheet, ByVal strTableNameIn As String, _
                            Optional ByVal displayError As Boolean = False) As Boolean
    
    On Error GoTo ErrTableDNE
    If Not WrkShtIn.ListObjects(strTableNameIn) Is Nothing Then TableExists = True
    Exit Function
    
ErrTableDNE:
    If displayError Then
        MsgBox "The table named " & strTableNameIn & " in Worksheet: " & _
                WrkShtIn.name & " either does not exist or it has been deleted. "
    End If
End Function

