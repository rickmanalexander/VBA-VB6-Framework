Attribute VB_Name = "WorkBookObject"
'@Folder("Framework.CommonMethods")
Option Explicit
Option Private Module

Public Function VisibleSheetCount(Optional wrkBook As Workbook) As Long

    Dim currentWorkBook As Workbook
    Set currentWorkBook = IIf(wrkBook Is Nothing, ThisWorkbook, wrkBook)
    
    Dim sht As Worksheet
    Dim counter As Long
    For Each sht In currentWorkBook.Worksheets
        If sht.Visible = Excel.XlSheetVisibility.xlSheetVisible Then counter = counter + 1
        
    Next
    
    VisibleSheetCount = counter
    
End Function


Public Function GetShtFromCodeName(ByVal strShtCodeName As String, ByRef WrkShtToReturn As Worksheet, _
                                   Optional ByVal displayError As Boolean = False) As Boolean
 
    Dim WrkSht As Worksheet
    For Each WrkSht In ThisWorkbook.Worksheets
         If UCase$(WrkSht.CodeName) = strShtCodeName Then
              Set WrkShtToReturn = WrkSht
              GetShtFromCodeName = True
              Exit Function
         End If
    Next
    
    If displayError Then
        If Not GetShtFromCodeName Then
            MsgBox "The WorkSheet Codenamed: " & strShtCodeName & " either does not exist or it has been deleted. "
        End If
    End If
   
End Function

Public Sub DeleteCreatedWrkBk(ByRef WrkBkIn As Workbook, ByVal boolTargetWasCreated As Boolean)
    
 Dim strWrkBkFullName As String
        
    'Check if the workBook object is instatiated
    If WrkBkIn Is Nothing Then Exit Sub
    
    strWrkBkFullName = WrkBkIn.FullName
    
    On Error GoTo CleanFail
    If Application.EnableEvents Then Application.EnableEvents = False
        With WrkBkIn
            'If this Workbook was created programatically
            If boolTargetWasCreated Then
                'if it has no name, was not provided a "SaveAs" Name, or was not
                'saved at (covers the case that an attempt was made to save the
                'file in the same location
                If Len(Trim$(.name)) = 0 Or .name = "False" Or Not .Saved Then
                    .Saved = True
                    .Close
                    Set WrkBkIn = Nothing
                Else
                    'remove readonly attribute, if set
                    SetAttr strWrkBkFullName, vbNormal
                    'avoids prompt asking if you want to save
                    .Saved = True
                    .Close SaveChanges:=False
                    'must set to nothing before deleting file
                    Set WrkBkIn = Nothing
                    'delete the file
                    Kill strWrkBkFullName
                End If
            End If
        End With
    Exit Sub
    
CleanExit:
    Exit Sub

CleanFail:
    Set WrkBkIn = Nothing
    Resume CleanExit
    
End Sub
