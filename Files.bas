Attribute VB_Name = "Files"
'@Folder("Framework.CommonMethods")

Option Explicit
Option Private Module

'**************************************************************************************************************************
'File Methods
'**************************************************************************************************************************
Public Function GetCheckMacroEnabledFilePath(ByVal workBookPath As String, ByVal initialFileName As String, _
                                             ByRef outFilePath As String, ByRef canceled As Boolean) As Boolean

    If canceled Then Exit Function

    outFilePath = Application.GetSaveAsFilename(initialFileName:=initialFileName, _
                                                FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
                                                Title:="Save As Macro-Enabled Workbook (.xlsm extension)")

    '"False" means they hit the cancel button on the
    'dialog
    canceled = (outFilePath = "False")

    If outFilePath = workBookPath Then

        canceled = (MsgBox("The Application already exists In this path. " & vbNewLine & vbNewLine & _
                           "Click retry to choose another path.", vbRetryCancel, _
                           "Error: Cannot Save") = vbCancel)

        'Recurse Function
        GetCheckMacroEnabledFilePath = GetCheckMacroEnabledFilePath(workBookPath, initialFileName, outFilePath, canceled)

    End If

    If outFilePath = "False" Then Exit Function
    
    If GetFileNameFromPath(outFilePath) <> initialFileName Then
        
            canceled = (MsgBox("You cannot change the name of this Application. " & vbNewLine & vbNewLine & _
                               "Click retry to save using the original name.", _
                               vbRetryCancel, "Error: Cannot Save") = vbCancel)

            'Recurse Function
           GetCheckMacroEnabledFilePath = GetCheckMacroEnabledFilePath(workBookPath, initialFileName, outFilePath, canceled)

    End If

    GetCheckMacroEnabledFilePath = ((Len(Trim$(outFilePath)) > 0) And (Not canceled))

End Function


Public Function GetFilePath(ByRef varfileType As Variant, ByVal DialogTitle As String, _
                            ByRef ReturnFilePath As String, Optional initialFileName As String) As Boolean
    
        With Application.FileDialog(msoFileDialogFilePicker)
            .Filters.Clear
            .Title = DialogTitle
            .initialFileName = initialFileName
            
            If VarType(varfileType) = (vbArray + vbVariant) Then
                .Filters.Add "File type", "*." & Join(varfileType, ", *.")
            End If
           
            'show the file picker dialog box
            If .Show <> 0 Then
                ReturnFilePath = .SelectedItems(1)
                GetFilePath = True
                Debug.Print ReturnFilePath
            Else
                ' MsgBox "You must select a folder or file path.", _
                        ' vbCritical, "Error: No Selection"
                GetFilePath = False
                Exit Function
            End If
            
        End With
    
End Function

Public Function GetFileNameFromPath(ByVal FilePath As String) As String
    'check if valid file path and not a folder paths
    If InStrRev(FilePath, ".") = 0 Then Exit Function
    GetFileNameFromPath = Mid(FilePath, InStrRev(FilePath, "\") + 1)
End Function

Public Function GetFileExtenstion(ByVal strFileName As String) As String
    GetFileExtenstion = Right$(strFileName, _
                             (Len(strFileName) - InStrRev(strFileName, ".")))
End Function

Public Function FileExists(ByVal FilePath As String) As Boolean

    With CreateObject("scripting.filesystemobject")
        FileExists = .FileExists(FilePath)
    End With
        
End Function


