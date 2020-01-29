Attribute VB_Name = "mVBProjectReferenceManagement"
Option Explicit
Option Private Module

Private Const ADODB_GUID As String = "{B691E011-1797-432E-907A-4D8C69339129}"
Private Const ADODB_MAJOR_VERSION As Long = 6
Private Const ADODB_MINOR_VERSION As Long = 1

Private Const MSCOREE_GUID As String = "{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}"
Private Const MSCOREE_MAJOR_VERSION As Long = 2
Private Const MSCOREE_MINOR_VERSION As Long = 4

Private Const MSCORLIB_GUID As String = "{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}"
Private Const MSCORLIB_MAJOR_VERSION As Long = 2
Private Const MSCORLIB_MINOR_VERSION As Long = 4

Public Sub CheckAndFixReferences()
    AddIfNotExists ADODB_GUID, ADODB_MAJOR_VERSION, ADODB_MINOR_VERSION
    AddIfNotExists MSCOREE_GUID, MSCOREE_MAJOR_VERSION, MSCOREE_MINOR_VERSION
    AddIfNotExists MSCORLIB_GUID, MSCORLIB_MAJOR_VERSION, MSCORLIB_MINOR_VERSION
End Sub

Private Sub AddIfNotExists(ByVal Guid As String, MajorVersion As Long, MinorVersion As Long)

    Const ErrorAlreadyExists As Long = 32813
    
    On Error GoTo ReferenceAlreadyExists
    ThisWorkbook.VBProject.References.AddFromGuid Guid, MajorVersion, MinorVersion

CleanExit:
    Exit Sub
    
ReferenceAlreadyExists:
    If Err.Number = ErrorAlreadyExists Then Resume CleanExit
End Sub


Private Sub GetReferences()

    Dim Ref As Object

    For Each Ref In ThisWorkbook.VBProject.References
        If Not Ref.BuiltIn Then PrintToImmediateWindow Ref
    Next

End Sub


Private Sub PrintToImmediateWindow(ByVal RefIn As Object)
    With RefIn
        Debug.Print "'--------------------------------------------------------------------------"
        Debug.Print "' Name:                  " + .Name
        Debug.Print "' BuiltIn:               " + CStr(.BuiltIn)
        Debug.Print "' GUID:                  " + .Guid
        Debug.Print "' Major Version Number:  " + CStr(.Major)
        Debug.Print "' Minor Version Number:  " + CStr(.Minor)
        Debug.Print "' Full Version Number:   " + CStr(.Major) + "." + CStr(.Minor)
        Debug.Print "' FullPath:              " + .FullPath
        Debug.Print "' Description:           " + .Description
        Debug.Print "'--------------------------------------------------------------------------"
    End With
End Sub


