VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbAutoDeriveParameterMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.AdoDbCommand.Parameters")
Option Explicit

Implements IAdoDbParameterMapper

Private Enum AdoDbDerivedCommandErrors
    DoesNotSupportParameterRetrieval
End Enum

Private Enum ParameterIndexIncrement
    commandText
    StoredProcedure
End Enum

Private Sub IAdoDbParameterMapper_SetParameters(ByRef commandReturn As ADODB.Command, ByRef parameterValues As Variant)
    
    If UBound(parameterValues) = -1 Then Exit Sub 'not allocated
        
    If commandReturn.Parameters.Count = 0 Then
        Err.Raise Number:=AdoDbDerivedCommandErrors.DoesNotSupportParameterRetrieval, _
                  Source:=TypeName(Me) & ".SetParameters", _
                  Description:="This Provider does not support parameter retrieval."
    End If
                                                                                                                                                                                            
    Select Case commandReturn.commandType

        Case adCmdStoredProc
            'Debug.Print Cmnd.Parameters.Count prints 1 b/c it includes '@RETURN_VALUE'
            'which is a default value
            If commandReturn.Parameters.Count > 1 Then ProcessParameters commandReturn, parameterValues, _
                                                                         ParameterIndexIncrement.StoredProcedure
                                                                                                                
        Case adCmdText
            ProcessParameters commandReturn, parameterValues, ParameterIndexIncrement.commandText
            
    End Select
           
End Sub

Private Sub ProcessParameters(ByRef commandReturn As ADODB.Command, ByRef parameterValues As Variant, ByVal increment As ParameterIndexIncrement)
    
    Dim i As Long
    Dim paramVal As Variant
    
    For i = LBound(parameterValues) To UBound(parameterValues)
        
        paramVal = parameterValues(i)
        
        If IsArray(paramVal) Then
            ProcessArrayParameter commandReturn, paramVal, increment
        
        Else
            'Explicitly set size to prevent error
            'as per the Note at: https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/refresh-method-ado?view=sql-server-2017
            SetVariableLengthProperties commandReturn.Parameters(i + increment), paramVal
                                                 
            commandReturn.Parameters(i + increment).value = paramVal  '.Parameters(i + 1) b/c of @RETURN_VALUE
                                                                  
                                                                      'mentioned above
        End If
        
    Next i
                
        
End Sub


Private Sub ProcessArrayParameter(ByRef commandReturn As ADODB.Command, _
                                  ByRef arrayParameter As Variant, _
                                  ByVal increment As ParameterIndexIncrement)

    AdoDbParameterHelper.ValidateArrayDimensions arrayParameter
    
    Dim j As Long
    Dim paramVal As Variant
    
    Select Case AdoDbParameterHelper.NumberOfArrayDimensions(arrayParameter)
        
        Case 1
            For j = LBound(arrayParameter) To UBound(arrayParameter)
                
                paramVal = arrayParameter(j)
                
                'Explicitly set size to prevent error
                'as per the Note at: https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/refresh-method-ado?view=sql-server-2017
                SetVariableLengthProperties commandReturn.Parameters(j + increment), paramVal
                                                     
                commandReturn.Parameters(j + increment).value = paramVal
                                                     
            Next j
        
        Case 2
            For j = LBound(arrayParameter, 1) To UBound(arrayParameter, 1)
                
                paramVal = arrayParameter(j, 1)
                
                'Explicitly set size to prevent error
                'as per the Note at: https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/refresh-method-ado?view=sql-server-2017
                SetVariableLengthProperties commandReturn.Parameters(j + increment), paramVal
                                                     
                commandReturn.Parameters(j + increment).value = paramVal
                                                     
            Next j
            
    End Select
        
End Sub
    
    
Private Sub SetVariableLengthProperties(ByRef parameter As ADODB.parameter, ByRef parameterValue As Variant)
            
    With parameter
        Select Case VarType(parameterValue)
                
            Case vbSingle
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
            
            Case vbDouble
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
                
            Case vbCurrency
                .precision = AdoDbParameterHelper.CalculatePrecision(parameterValue)
                .numericScale = AdoDbParameterHelper.CalculateNumericScale(parameterValue)
                
            Case vbString
                .size = Len(parameterValue)
            
        End Select

    End With
            
End Sub

