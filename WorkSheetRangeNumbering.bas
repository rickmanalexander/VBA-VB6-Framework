Attribute VB_Name = "WorkSheetRangeNumbering"
'@Folder("Framework.CommonMethods")

Option Explicit
Option Private Module

Public Function FindLastRowInWrkSht(ByRef WrkShtIn As Worksheet) As Long

    With WrkShtIn
        If Application.CountA(.Cells) <> 0 Then
            FindLastRowInWrkSht = .Cells.Find(What:="*", _
                          After:=.Range("A1"), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious, _
                          MatchCase:=False).Row
        End If
    End With
    
End Function

Public Function FindLastColumnInWrkSht(ByRef WrkShtIn As Worksheet) As Long

    With WrkShtIn
        If Application.CountA(.Cells) <> 0 Then
            FindLastColumnInWrkSht = .Cells.Find(What:="*", _
                          After:=.Range("A1"), _
                          Lookat:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious, _
                          MatchCase:=False).Column
        End If
    End With
    
End Function

Public Function FindLastRowInAColumn(ByRef WrkShtIn As Worksheet, ByVal strColToFindLastRow As String)
    FindLastRowInAColumn = WrkShtIn.Cells(WrkShtIn.Rows.Count, strColToFindLastRow).End(xlUp).Row
End Function

Public Function FindLastRowInCurrentRegion(ByRef WrkShtIn As Worksheet, ByVal strRangeAddress As String)
    FindLastRowInCurrentRegion = WrkShtIn.Range(strRangeAddress).CurrentRegion _
                                .Rows(WrkShtIn.Range(strRangeAddress).CurrentRegion.Rows.Count).Row
End Function
        
Public Function FindLastColumnAddress(ByRef WrkShtIn As Worksheet) As String
        FindLastColumnAddress = GetColumnLetterFromNumber(WrkShtIn, FindLastColumnInWrkSht(WrkShtIn))
End Function

Public Function GetColumnLetterFromNumber(ByRef WrkShtIn As Worksheet, ByVal lngColumnNumberIn As Long) As String
    GetColumnLetterFromNumber = Split(WrkShtIn.Cells(1, lngColumnNumberIn).Address, "$")(1)
End Function

Public Function GetColumnNumberFromLetter(ByRef WrkShtIn As Worksheet, ByVal strColumnLetterIn As String) As String
    strColumnLetterIn = StripNumbers(strColumnLetterIn)
    GetColumnNumberFromLetter = WrkShtIn.Range(strColumnLetterIn & 1).Column
End Function

