VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbNamedParameterMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.AdoDbCommand.Parameters")
Option Explicit

'Implements IAdoDbParameterMapper
'
'Private Sub IAdoDbParameterMapper_SetParameters(ByRef commandReturn As ADODB.Command, ByRef parameterValues As Variant)
'
'    If UBound(parameterValues) = -1 Then Exit Sub 'not allocated
'
'    Dim i As Long
'    For i = LBound(parameterValues) To UBound(parameterValues)
'
'        If parameterValues(i) Is ADODB.parameter Then
'            commandReturn.Parameters.Append parameterValues(i)
'
'        ElseIf IsArray(parameterValues(i)) Then
'            ProcessArrayParameter commandReturn, parameterValues(i)
'
'        ElseIf parameterValues(i) Is ADODB.Parameters Then
'            ProcessParametersCollection commandReturn, paramVal
'
'        End If
'
'    Next i
'
'End Sub
'
'Private Sub ProcessArrayParameter(ByRef commandReturn As ADODB.Command, ByRef arrayParameter As Variant)
'
'    AdoDbParameterHelper.ValidateArrayDimensions arrayParameter
'
'    Dim j As Long
'    Dim paramVal As Variant
'
'    Select Case AdoDbParameterHelper.NumberOfArrayDimensions(arrayParameter)
'
'        Case 1
'            For j = LBound(arrayParameter) To UBound(arrayParameter)
'                If arrayParameter(i) Is ADODB.parameter Then
'                    commandReturn.Parameters.Append arrayParameter(i)
'
'                End If
'
'            Next j
'
'        Case 2
'            For j = LBound(arrayParameter, 1) To UBound(arrayParameter, 1)
'                If arrayParameter(j, 1) Is ADODB.parameter Then
'                    commandReturn.Parameters.Append arrayParameter(j, 1)
'
'                End If
'
'            Next j
'
'    End Select
'
'
'End Sub
'
'
'Private Sub ProcessParametersCollection(ByRef commandReturn As ADODB.Command, ByRef parameterCollection As ADODB.Parameters)
'
'    Dim param As ADODB.parameter
'    For Each param In parameterCollection
'        commandReturn.Parameters.Append param
'    Next
'
'End Sub
'
'
