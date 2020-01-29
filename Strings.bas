Attribute VB_Name = "Strings"
'@Folder("Framework.CommonMethods")

Option Explicit
Option Private Module

Public Function StripNumbers(strInPut As String) As String
    With CreateObject("VBScript.RegExp")
        .Global = True
        .pattern = "\d+"
        StripNumbers = .Replace(strInPut, vbNullString)
    End With
End Function

Public Function RegExMatchesCount(ByVal patternValue As String, ByRef value As String) As Long
    
    With CreateObject("VBScript.RegExp")
        .pattern = patternValue
        .Global = True
        'If Not .test(value) Then Exit Function
        RegExMatchesCount = .Execute(value).Count
    End With
    
End Function

Public Function IsEmptyString(ByVal value As String) As Boolean
    
    RemoveAllWhiteSpace value
    IsEmptyString = (Len(value) = 0)
    
End Function

Public Function RemoveAllWhiteSpace(ByRef varStringIn As Variant, _
                                    Optional ByRef ObjRegExpIn As Object) As String

    'Create if not instantiated
    If ObjRegExpIn Is Nothing Then Set ObjRegExpIn = CreateObject("VBScript.RegExp")
    
    With ObjRegExpIn
        .pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveAllWhiteSpace = CStr(.Replace(varStringIn, vbNullString))
    End With
    
End Function

Public Function RemoveNCharactersFromEnd(ByVal stringIn As String, ByVal numberOfCharacters As Long) As String

    stringIn = Trim$(stringIn)
    
    RemoveNCharactersFromEnd = Left$(stringIn, Len(stringIn) - numberOfCharacters)
    
End Function

Public Function RemoveNCharactersFromBegining(ByVal stringIn As String, ByVal numberOfCharacters As Long) As String

    stringIn = Trim$(stringIn)
    
    RemoveNCharactersFromBegining = Right$(stringIn, Len(stringIn) - numberOfCharacters)
    
End Function

Public Function PrintF(ByVal text As String, ParamArray Args() As Variant) As String

    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim formatString As String
    Dim argumentLength As Long
    
    Dim returnString As String
    returnString = text
    
    For i = LBound(Args) To UBound(Args)
    
        argumentLength = Len(CStr(i))
        
        startPos = InStr(returnString, "{" & CStr(i) & ":")
        
        If startPos > 0 Then
            endPos = InStr(startPos + 1, returnString, "}")
            formatString = Mid(returnString, startPos + 2 + argumentLength, endPos - (startPos + 2 + argumentLength))
            returnString = Mid(returnString, 1, startPos - 1) & Format(Args(i), formatString) & Mid(returnString, endPos + 1)
        
        Else
            returnString = Replace(returnString, "{" & CStr(i) & "}", Args(i))
            
        End If
    Next i

    PrintF = returnString

End Function


Public Function ExtractString(ByRef value As String, ByVal firstDelimiter As String, _
                              ByVal secondDelimiter As String) As String
    
    Dim firstPosition As Integer
    firstPosition = InStr(value, firstDelimiter)

    Dim secondPosition As Integer
    secondPosition = InStr(value, secondDelimiter)
    
    If firstPosition = 0 Or secondPosition = 0 Then Exit Function
    
    Dim extractedString As String
    extractedString = Mid$(value, firstPosition + 1, secondPosition - firstPosition - 1)
    
    If LenB(extractedString) > 0 Then ExtractString = extractedString

End Function

