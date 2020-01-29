Attribute VB_Name = "Arrays"
'@Folder("Framework.CommonMethods")

Option Explicit
Option Private Module

Private Enum ArrayErrors
    WorkSheetObjectIsNothing = vbObjectError + 1024
    InValidRangeAddress
    EmptyRange
    ArrayNotAllocated
    InValidDimensions
End Enum


'**************************************************************************************************************
'Array Functions / Procedures
'       Function: RangeToArray
'       Procedure: ArrayToRange
'
'   '*******************************
'    Helper Functions
'   '*******************************
'                   RangeToArray
'                   ArrayToRange
'**************************************************************************************************************
Public Function RangeToArray(ByRef returnArray As Variant, ByRef wrkSheet As Worksheet, _
                             ByVal rangeBeginAddress As String, ByVal rangeEndAddress As String, _
                             ByVal keepRangeFormat As Boolean) As Boolean
    
    If wrkSheet Is Nothing Then ThrowError ArrayErrors.WorkSheetObjectIsNothing, "Framework.Arrays.RangeToArray"
    
    If GetColumnNumberFromLetter(wrkSheet, rangeBeginAddress) > _
       GetColumnNumberFromLetter(wrkSheet, rangeEndAddress) Then _
       ThrowError ArrayErrors.InValidRangeAddress, "Framework.Arrays.RangeToArray"
       
    If Application.CountA(wrkSheet.Cells) = 0 Then _
        ThrowError ArrayErrors.EmptyRange, "Framework.Arrays.RangeToArray"
        
        With wrkSheet
            If keepRangeFormat Then
                returnArray = .Range(rangeBeginAddress & ":" & rangeEndAddress).value
                
            Else
                returnArray = .Range(rangeBeginAddress & ":" & rangeEndAddress).Value2
                
            End If
            
            RangeToArray = True
            
        End With

        If Not IsArray(returnArray) Then _
            ThrowError ArrayErrors.ArrayNotAllocated, "Framework.Arrays.RangeToArray"
    
End Function

'**************************************************************************************************************

Public Sub ArrayToRange(ByRef arryIn As Variant, ByRef wrkSheet As Worksheet, _
                        ByVal rangeOut As String, Optional singleDimensionArrayToRow As Boolean = False)
 
    Dim firstDimUpperBound As Long, secondDimUpperBound As Long
    
    On Error GoTo CleanFail
    If NumberOfArrayDimensions(arryIn) = 2 Then
        If IsArrayZeroBased(arryIn) Then
            firstDimUpperBound = UBound(arryIn, 1) + 1
            secondDimUpperBound = UBound(arryIn, 2) + 1
        
            wrkSheet.Range(rangeOut).Resize(firstDimUpperBound, _
            secondDimUpperBound).Value2 = arryIn
            
        Else
            firstDimUpperBound = UBound(arryIn, 1)
            secondDimUpperBound = UBound(arryIn, 2)
            
            wrkSheet.Range(rangeOut).Resize(firstDimUpperBound, _
            secondDimUpperBound).Value2 = arryIn
            
        End If
        
    Else
        If IsArrayZeroBased(arryIn) Then
            If singleDimensionArrayToRow Then
                firstDimUpperBound = UBound(arryIn, 1) + 1
                '1-D Arry to 1 Row
                wrkSheet.Range(rangeOut).Resize(1, firstDimUpperBound).Value2 = To2dArray(arryIn, True)
                
            Else
                firstDimUpperBound = UBound(arryIn) + 1
                '1-D Arry to 1 column
                wrkSheet.Range(rangeOut).Resize(UBound(arryIn)).Value2 = To2dArray(arryIn, False)
                
            End If
            
        Else
            If singleDimensionArrayToRow Then
                firstDimUpperBound = UBound(arryIn, 1)
                '1-D Arry to 1 Row
                wrkSheet.Range(rangeOut).Resize(1, firstDimUpperBound).Value2 = To2dArray(arryIn, True)
                
            Else
                firstDimUpperBound = UBound(arryIn)
                '1-D Arry to 1 column
                wrkSheet.Range(rangeOut).Resize(firstDimUpperBound).Value2 = To2dArray(arryIn, False)
                
            End If
            
        End If
        
    End If
    
CleanExit:
    Exit Sub
    
CleanFail:
    Resume CleanExit

End Sub

Public Function IsArrayZeroBased(arryIn As Variant) As Boolean
    
    If Not IsArrayAllocated(arryIn) Then Exit Function
    
    If NumberOfArrayDimensions(arryIn) = 1 Then
        If LBound(arryIn, 1) = 0 Then IsArrayZeroBased = True
        
    Else
        If LBound(arryIn, 1) = 0 And UBound(arryIn, 2) = 0 Then IsArrayZeroBased = True
        
    End If

End Function

Public Function NumberOfArrayDimensions(variantArray As Variant) As Integer

    Dim index As Long, upperBound As Long

        On Error Resume Next
        Err.Clear
        Do
            index = index + 1
            upperBound = UBound(variantArray, index)
        Loop Until Err.Number <> 0

    NumberOfArrayDimensions = index - 1

End Function

Public Function IsArrayAllocated(arryIn As Variant) As Boolean
 
    On Error Resume Next
    If Not IsArray(arryIn) Then Exit Function

    If Err.Number = 0 Then
        If LBound(arryIn) <= UBound(arryIn) Then IsArrayAllocated = True
    End If

End Function

Public Function IsMultiColumnArray(variantArray As Variant) As Boolean

    On Error Resume Next
    Err.Clear

    Dim value As Variant
    value = variantArray(LBound(variantArray), 2)

    IsMultiColumnArray = (Err.Number = 0)

End Function


Public Function TransPose2dArray(ByRef sourceArray As Variant) As Variant()
    
    Dim i As Long, j As Long
    
    Dim outArray As Variant
    ReDim outArray(LBound(sourceArray, 2) To UBound(sourceArray, 2), _
                   LBound(sourceArray, 1) To UBound(sourceArray, 1))
    
    For j = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        For i = LBound(sourceArray, 1) To UBound(sourceArray, 1)
            outArray(j, i) = sourceArray(i, j)
        Next i
    Next j
    
    TransPose2dArray = outArray
    
End Function

Public Function To2dArray(ByRef singleDimensionArray As Variant, ByVal horizontalOrientation As Boolean) As Variant()
    
    If NumberOfArrayDimensions(singleDimensionArray) > 1 Then ThrowError ArrayErrors.InValidDimensions, "Framework.Arrays.To2dArray"
    
    Dim outArray As Variant
    
    Dim lowerBound As Long: lowerBound = LBound(singleDimensionArray)
    
    If horizontalOrientation Then
        ReDim outArray(lowerBound To lowerBound, LBound(singleDimensionArray) To UBound(singleDimensionArray))
        
        Dim j As Long
        For j = LBound(singleDimensionArray) To UBound(singleDimensionArray)
            outArray(lowerBound, j) = singleDimensionArray(j)
        Next j
        
    Else
        ReDim outArray(LBound(singleDimensionArray) To UBound(singleDimensionArray), lowerBound To lowerBound)
        
        Dim i As Long
        For i = LBound(singleDimensionArray) To UBound(singleDimensionArray)
            outArray(i, lowerBound) = singleDimensionArray(i)
        Next i
        
    End If
    
    To2dArray = outArray
    
End Function

    
Private Sub ThrowError(ByVal errorNumber As ArrayErrors, ByVal qaulifiedMethodName As String)

    Select Case errorNumber
        Case ArrayErrors.WorkSheetObjectIsNothing
            Err.Raise ArrayErrors.WorkSheetObjectIsNothing, qaulifiedMethodName, "The worksheet passed is not valid."
            
        Case ArrayErrors.InValidRangeAddress
            Err.Raise ArrayErrors.InValidRangeAddress, qaulifiedMethodName, "The range start address is greater than the range end address."
            
        Case ArrayErrors.EmptyRange
            Err.Raise ArrayErrors.EmptyRange, qaulifiedMethodName, "The range passed is empty"
            
            
        Case ArrayErrors.ArrayNotAllocated
            Err.Raise ArrayErrors.ArrayNotAllocated, qaulifiedMethodName, "The return array was not allocated"
            
        Case ArrayErrors.InValidDimensions
            Err.Raise ArrayErrors.InValidDimensions, qaulifiedMethodName, "The array passed was not a single dimension."
    
    End Select
    
End Sub


