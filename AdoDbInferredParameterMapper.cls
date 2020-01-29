VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbInferredParameterMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.AdoDbCommand.Parameters")
Option Explicit

Implements IAdoDbParameterMapper

Private Sub IAdoDbParameterMapper_SetParameters(ByRef commandReturn As ADODB.Command, ByRef parameterValues As Variant)

    Dim i As Long
    Dim paramVal As Variant
    
    If UBound(parameterValues) = -1 Then Exit Sub 'not allocated
    
    For i = LBound(parameterValues) To UBound(parameterValues)
        paramVal = parameterValues(i)
        
        If IsArray(paramVal) Then
            ProcessArrayParameter commandReturn, paramVal
            
        Else
            commandReturn.Parameters.Append ToAdoInputParameter(paramVal)
            
        End If
        
    Next i

End Sub

Private Sub ProcessArrayParameter(ByRef commandReturn As ADODB.Command, ByRef arrayParameter As Variant)

    AdoDbParameterHelper.ValidateArrayDimensions arrayParameter
    
    Dim j As Long
    Dim paramVal As Variant
    
    Select Case AdoDbParameterHelper.NumberOfArrayDimensions(arrayParameter)
        
        Case 1
            For j = LBound(arrayParameter) To UBound(arrayParameter)
                paramVal = arrayParameter(j)
                
                commandReturn.Parameters.Append ToAdoInputParameter(paramVal)
                                                     
            Next j
        
        Case 2
            For j = LBound(arrayParameter, 1) To UBound(arrayParameter, 1)
                paramVal = arrayParameter(j, 1)
                
                commandReturn.Parameters.Append ToAdoInputParameter(paramVal)
                                                     
            Next j
            
    End Select
        

End Sub

Private Function ToAdoInputParameter(ByVal parameterValue As Variant) As ADODB.parameter

    Dim resultParameter As ADODB.parameter
    Set resultParameter = New ADODB.parameter

    With resultParameter
        Select Case VarType(parameterValue)
        
            Case vbInteger
                .Type = adInteger

            Case vbLong
                .Type = adInteger
                
            Case vbSingle
                .Type = adSingle
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
            
            Case vbDouble
                .Type = adDouble
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
            
            Case vbDate
                .Type = adDate
                
            Case vbCurrency
                .Type = adCurrency
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
                
            Case vbString
                .Type = adVarChar
                .size = Len(parameterValue)
            
            Case vbBoolean
                .Type = adBoolean
            
        End Select

        .direction = ADODB.ParameterDirectionEnum.adParamInput
        .value = parameterValue
    End With
        
    Set ToAdoInputParameter = resultParameter

End Function

