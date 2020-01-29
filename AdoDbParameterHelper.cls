VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbParameterHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.AdoDbCommand.Parameters")

Option Explicit
'@PredeclaredId

Public Enum AdoDbParameterHelperErrors
    ParameterIsArrayOfThreeOrMoreDimensions = vbObjectError + 1024
    ParameterIsMultiColumnArray
    NonDefaultInstanceCreated
End Enum

Public Function CalculatePrecision(ByVal value As Variant) As Byte
    CalculatePrecision = CByte(Len(Replace(CStr(value), ".", vbNullString)))
End Function

Public Function CalculateNumericScale(ByVal value As Variant) As Byte
    CalculateNumericScale = CByte(Len(Split(CStr(value), ".")(1)))
End Function


Public Sub ValidateArrayDimensions(ByRef variantArray As Variant)

    Dim dimensions As Long
    dimensions = NumberOfArrayDimensions(variantArray)
    Select Case dimensions
        Case Is > 2
            Err.Raise Number:=AdoDbParameterHelperErrors.ParameterIsArrayOfThreeOrMoreDimensions, _
                      Source:=TypeName(Me) & "." & "ProcessArrayParameter", _
                      Description:="Current parameter is an array of 3 or mor dimensions"
        Case Is = 2
            If IsMultiColumnArray(variantArray) Then Err.Raise Number:=AdoDbParameterHelperErrors.ParameterIsMultiColumnArray, _
                                                               Source:=TypeName(Me) & "." & "ProcessArrayParameter", _
                                                               Description:="Current parameter is a 2-dimensional array, with more, " & _
                                                                            "than one column"

    End Select

End Sub


Public Function NumberOfArrayDimensions(ByRef variantArray As Variant) As Long

    Dim index As Long, upperBound As Long

        On Error Resume Next
        Err.Clear
        Do
            index = index + 1
            upperBound = UBound(variantArray, index)
        Loop Until Err.Number <> 0

    NumberOfArrayDimensions = index - 1

End Function

Public Function IsMultiColumnArray(ByRef variantArray As Variant) As Boolean

    On Error Resume Next
    Err.Clear

    Dim value As Variant
    value = variantArray(LBound(variantArray), 2)

    IsMultiColumnArray = (Err.Number = 0)

End Function


Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is AdoDbParameterHelper
End Property


Private Sub Class_Initialize()
    If Not IsDefaultInstance Then Err.Raise Number:=AdoDbParameterHelperErrors.NonDefaultInstanceCreated, _
                                            Source:=TypeName(Me), _
                                            Description:="Cannot instantiate a static class"
End Sub
