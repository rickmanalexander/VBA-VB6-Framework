VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess")
Option Explicit

Private Type TAdoDbDataAccess
    DeriveParameterDirection As Boolean
    CommandTimeout As Long
    OuputParameters As Collection
    ADOErrors As ADODB.Errors
    HasError As Boolean
End Type

Private this As TAdoDbDataAccess

Public Property Get DeriveParameterDirection() As Boolean
    DeriveParameterDirection = this.DeriveParameterDirection
End Property

Public Property Let DeriveParameterDirection(ByVal value As Boolean)
    this.DeriveParameterDirection = value
End Property

Public Property Get CommandTimeout() As Long
    CommandTimeout = this.CommandTimeout
End Property

Public Property Let CommandTimeout(ByVal value As Long)
    this.CommandTimeout = value
End Property

Public Property Get OuputParameters() As Collection
    Set OuputParameters = this.OuputParameters
End Property

Public Property Get Errors() As ADODB.Errors
    Set Errors = this.ADOErrors
End Property

Public Property Get HasError() As Boolean
    HasError = this.HasError
End Property

Private Sub Class_Terminate()
    With this
        .CommandTimeout = 0
        .DeriveParameterDirection = Empty
        Set .OuputParameters = Nothing
        Set .ADOErrors = Nothing
        .HasError = Empty
    End With
End Sub

Public Function ExecuteQuery(ByRef connection As ADODB.connection, _
                             ByVal commandText As String, _
                             ByVal commandType As ADODB.CommandTypeEnum, _
                             ByVal cursorType As ADODB.CursorTypeEnum, _
                             ByVal lockType As ADODB.LockTypeEnum, _
                             ParamArray parameterValues() As Variant) As ADODB.recordset

    ValidateConnection connection.Errors
        
    On Error GoTo CleanFail
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(connection, commandText, commandType, CVar(parameterValues)) 'must convert paramarray to
                                                                                          'a variant in order to pass                                                                         'to another function
    'Note:  When used on a client-side Recordset object,
    '       the cursorType property can be set only to adOpenStatic.
    Set ExecuteQuery = New ADODB.recordset
        ExecuteQuery.cursorType = cursorType
        ExecuteQuery.lockType = lockType
    
    Set ExecuteQuery = cmnd.Execute()
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    Exit Function
    
CleanFail:
    PopulateADOErrorObject connection.Errors
    Resume CleanExit
    
End Function

Public Function ExecuteQueryAsync(ByRef connection As ADODB.connection, _
                                  ByVal commandText As String, _
                                  ByVal commandType As ADODB.CommandTypeEnum, _
                                  ByVal cursorType As ADODB.CursorTypeEnum, _
                                  ByVal lockType As ADODB.LockTypeEnum, _
                                  ParamArray parameterValues() As Variant) As ADODB.recordset

    ValidateConnection connection.Errors
        
    On Error GoTo CleanFail
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(connection, commandText, commandType, CVar(parameterValues)) 'must convert paramarray to
                                                                                          'a variant in order to pass                                                                         'to another function
    'Note:  When used on a client-side Recordset object,
    '       the cursorType property can be set only to adOpenStatic.
    Set ExecuteQueryAsync = New ADODB.recordset
        ExecuteQueryAsync.cursorType = cursorType
        ExecuteQueryAsync.lockType = lockType
    
    Set ExecuteQueryAsync = cmnd.Execute(Options:=ExecuteOptionEnum.adAsyncFetch)   'Async specified here
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    Exit Function
    
CleanFail:
    PopulateADOErrorObject connection.Errors
    Resume CleanExit
    
End Function

Public Function QuickExecuteQuery(ByRef ConnectionString As String, _
                                  ByVal cursorLocation As ADODB.CursorLocationEnum, _
                                  ByVal commandText As String, _
                                  ByVal commandType As ADODB.CommandTypeEnum, _
                                  ParamArray parameterValues() As Variant) As ADODB.recordset
                                              
    On Error GoTo CleanFail
    Dim CurrentConnection As ADODB.connection
    Set CurrentConnection = CreateConnection(ConnectionString, cursorLocation)
    
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(CurrentConnection, commandText, commandType, CVar(parameterValues)) 'must convert paramarray to
                                                                                                 'a variant in order to pass
                                                                                                 'to another function
    Set QuickExecuteQuery = New ADODB.recordset
    With QuickExecuteQuery
        .cursorType = adOpenStatic          'Must use this cursorType and this lockType to work with a disconnected recordset
        .lockType = adLockBatchOptimistic
        .Open cmnd
        
        'disconnect the recordset
        Set .ActiveConnection = Nothing
    End With

    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    If Not CurrentConnection Is Nothing Then: If (CurrentConnection.State And adStateOpen) = adStateOpen Then CurrentConnection.Close
    Set CurrentConnection = Nothing
    Exit Function
    
CleanFail:
    PopulateADOErrorObject CurrentConnection.Errors
    Resume CleanExit

End Function


Public Function QuickExecuteQueryAsync(ByRef ConnectionString As String, _
                                       ByVal cursorLocation As ADODB.CursorLocationEnum, _
                                       ByVal commandText As String, _
                                       ByVal commandType As ADODB.CommandTypeEnum, _
                                       ParamArray parameterValues() As Variant) As ADODB.recordset
                                              
    On Error GoTo CleanFail
    Dim CurrentConnection As ADODB.connection
    Set CurrentConnection = CreateConnection(ConnectionString, cursorLocation)
    
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(CurrentConnection, commandText, commandType, CVar(parameterValues)) 'must convert paramarray to
                                                                                                 'a variant in order to pass
                                                                                                 'to another function
    Set QuickExecuteQueryAsync = New ADODB.recordset
    With QuickExecuteQueryAsync
        .cursorType = adOpenStatic          'Must use this cursorType and this lockType to work with a disconnected recordset
        .lockType = adLockBatchOptimistic
        .Open cmnd, , , , Options:=ExecuteOptionEnum.adAsyncFetch   'Async specified here
        
        'disconnect the recordset
        Set .ActiveConnection = Nothing
    End With

    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    If Not CurrentConnection Is Nothing Then: If (CurrentConnection.State And adStateOpen) = adStateOpen Then CurrentConnection.Close
    Set CurrentConnection = Nothing
    Exit Function
    
CleanFail:
    PopulateADOErrorObject CurrentConnection.Errors
    Resume CleanExit

End Function

Public Function ExecuteNonQuery(ByRef connection As ADODB.connection, _
                                ByVal commandText As String, _
                                ByVal commandType As ADODB.CommandTypeEnum, _
                                ByRef returnRecordsAffected As Long, _
                                ParamArray parameterValues() As Variant) As Boolean

    ValidateConnection connection.Errors
    
    On Error GoTo CleanFail
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(connection, commandText, commandType, CVar(parameterValues))    'must convert paramarray to
                                                                                             'a variant in order to pass
                                                                                             'to another function
    cmnd.Execute RecordsAffected:=returnRecordsAffected, Options:=ExecuteOptionEnum.adExecuteNoRecords
    returnRecordsAffected = IIf((returnRecordsAffected = -1), 0, returnRecordsAffected)
    ExecuteNonQuery = True
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear

CleanExit:
    Set cmnd = Nothing
    Exit Function

CleanFail:
    PopulateADOErrorObject connection.Errors
    Resume CleanExit
    
End Function

Public Function ExecuteNonQueryAsync(ByRef connection As ADODB.connection, _
                                     ByVal commandText As String, _
                                     ByVal commandType As ADODB.CommandTypeEnum, _
                                     ByRef returnRecordsAffected As Long, _
                                     ParamArray parameterValues() As Variant) As Boolean

    ValidateConnection connection.Errors
    
    On Error GoTo CleanFail
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(connection, commandText, commandType, CVar(parameterValues))    'must convert paramarray to
                                                                                             'a variant in order to pass
                                                                                             'to another function
    cmnd.Execute RecordsAffected:=returnRecordsAffected, Options:=ExecuteOptionEnum.adExecuteNoRecords And ExecuteOptionEnum.adAsyncExecute
    returnRecordsAffected = IIf((returnRecordsAffected = -1), 0, returnRecordsAffected)
    ExecuteNonQueryAsync = True
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear

CleanExit:
    Set cmnd = Nothing
    Exit Function

CleanFail:
    PopulateADOErrorObject connection.Errors
    Resume CleanExit
    
End Function

'For the returnRecordsAffected to be populated the stored procedure, CANNOT Use: "Set NoCount On"
'in stored procedure
Public Function QuickExecuteNonQuery(ByVal ConnectionString As String, _
                                     ByVal commandText As String, _
                                     ByVal commandType As ADODB.CommandTypeEnum, _
                                     ByRef returnRecordsAffected As Long, _
                                     ParamArray parameterValues() As Variant) As Boolean
        
    On Error GoTo CleanFail
    Dim CurrentConnection As ADODB.connection
    Set CurrentConnection = CreateConnection(ConnectionString, adUseServer)
    
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(CurrentConnection, commandText, commandType, CVar(parameterValues))    'must convert paramarray to
                                                                                                    'a variant in order to pass
                                                                                                    'to another function
    cmnd.Execute RecordsAffected:=returnRecordsAffected, Options:=ExecuteOptionEnum.adExecuteNoRecords
    returnRecordsAffected = IIf((returnRecordsAffected = -1), 0, returnRecordsAffected)
    QuickExecuteNonQuery = True
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    If Not CurrentConnection Is Nothing Then: If (CurrentConnection.State And adStateOpen) = adStateOpen Then CurrentConnection.Close
    Set CurrentConnection = Nothing
    Exit Function

CleanFail:
    PopulateADOErrorObject CurrentConnection.Errors
    Resume CleanExit
    
End Function

Public Function QuickExecuteNonQueryAsync(ByVal ConnectionString As String, _
                                          ByVal commandText As String, _
                                          ByVal commandType As ADODB.CommandTypeEnum, _
                                          ByRef returnRecordsAffected As Long, _
                                          ParamArray parameterValues() As Variant) As Boolean
        
    On Error GoTo CleanFail
    Dim CurrentConnection As ADODB.connection
    Set CurrentConnection = CreateConnection(ConnectionString, adUseServer)
    
    Dim cmnd As ADODB.Command
    Set cmnd = CreateCommand(CurrentConnection, commandText, commandType, CVar(parameterValues))    'must convert paramarray to
                                                                                                    'a variant in order to pass
                                                                                                    'to another function
    cmnd.Execute RecordsAffected:=returnRecordsAffected, Options:=ExecuteOptionEnum.adExecuteNoRecords And ExecuteOptionEnum.adAsyncExecute
    returnRecordsAffected = IIf((returnRecordsAffected = -1), 0, returnRecordsAffected)
    QuickExecuteNonQueryAsync = True
    
    'if successful
    If Not this.ADOErrors Is Nothing Then this.ADOErrors.Clear
        
CleanExit:
    Set cmnd = Nothing
    If Not CurrentConnection Is Nothing Then: If (CurrentConnection.State And adStateOpen) = adStateOpen Then CurrentConnection.Close
    Set CurrentConnection = Nothing
    Exit Function

CleanFail:
    PopulateADOErrorObject CurrentConnection.Errors
    Resume CleanExit
    
End Function

Public Function CreateConnection(ByRef ConnectionString As String, ByVal cursorLocation As ADODB.CursorLocationEnum) As ADODB.connection

    On Error GoTo CleanFail
    Set CreateConnection = New ADODB.connection
        CreateConnection.cursorLocation = cursorLocation
        CreateConnection.Open ConnectionString

CleanExit:
    Exit Function

CleanFail:
    PopulateADOErrorObject CreateConnection.Errors
    Resume CleanExit

End Function

Private Function CreateCommand(ByRef connection As ADODB.connection, _
                               ByVal commandText As String, _
                               ByVal commandType As ADODB.CommandTypeEnum, _
                               ByRef parameterValues As Variant) As ADODB.Command

    Set CreateCommand = New ADODB.Command
    With CreateCommand
        .ActiveConnection = connection
        .commandText = commandText
        .CommandTimeout = Me.CommandTimeout '0
    End With
        
    Dim commandBuilder As IAdoDbParameterMapper
    If Me.DeriveParameterDirection Then
        Set commandBuilder = New AdoDbAutoDeriveParameterMapper
        
        CreateCommand.commandType = commandType         'When set before accessing the Parameters Collection,
                                                        'Parameters.Refresh is impilicitly called
                                                        
        commandBuilder.SetParameters CreateCommand, parameterValues
        
        PopulateOutPutParameters CreateCommand.Parameters
        
    Else
        Set commandBuilder = New AdoDbInferredParameterMapper
        
        commandBuilder.SetParameters CreateCommand, parameterValues
        
        CreateCommand.commandType = commandType

    End If
                
End Function

Private Sub ValidateConnection(ByRef ConnectionErrors As ADODB.Errors)

    If ConnectionErrors.Count > 0 Then
    
        If Not this.HasError Then PopulateADOErrorObject ConnectionErrors
        
        Dim ADOError As ADODB.Error
        Set ADOError = GetError(ConnectionErrors, ConnectionErrors.Count - 1) 'Note: 0 based collection
        
        Err.Raise ADOError.Number, ADOError.Source, ADOError.Description, ADOError.HelpFile, ADOError.HelpContext
        
    End If
    
End Sub

Private Sub PopulateADOErrorObject(ByVal ConnectionErrors As ADODB.Errors)
        
    If ConnectionErrors.Count = 0 Then Exit Sub
    
    this.HasError = True
    
    Set this.ADOErrors = ConnectionErrors

End Sub

Public Function ErrorsToString() As String

    Dim ADOError As ADODB.Error
    Dim i As Long
    Dim errorMsg As String
        
    For Each ADOError In this.ADOErrors
        i = i + 1
        With ADOError
            errorMsg = errorMsg & "Count: " & vbTab & i & vbNewLine
            errorMsg = errorMsg & "ADO Error Number: " & vbTab & CStr(.Number) & vbNewLine
            errorMsg = errorMsg & "Description: " & vbTab & .Description & vbNewLine
            errorMsg = errorMsg & "Source: " & vbTab & .Source & vbNewLine
            errorMsg = errorMsg & "NativeError: " & vbTab & CStr(.NativeError) & vbNewLine
            errorMsg = errorMsg & "HelpFile: " & vbTab & .HelpFile & vbNewLine
            errorMsg = errorMsg & "HelpContext: " & vbTab & CStr(.HelpContext) & vbNewLine
            errorMsg = errorMsg & "SQLState: " & vbTab & .SqlState & vbNewLine
        End With
    Next

    ErrorsToString = errorMsg & vbNewLine

End Function

Public Function GetError(ByRef ADOErrors As ADODB.Errors, ByVal index As Variant) As ADODB.Error
    Set GetError = IIf(ADOErrors.Count = 0, Nothing, ADOErrors.Item(index))
End Function

Private Sub PopulateOutPutParameters(ByRef Parameters As ADODB.Parameters)

    Set this.OuputParameters = New Collection

    Dim param As ADODB.parameter
    For Each param In Parameters
        Select Case param.direction
            Case ADODB.ParameterDirectionEnum.adParamInputOutput
                this.OuputParameters.Add param.value

            Case ADODB.ParameterDirectionEnum.adParamOutput
                this.OuputParameters.Add param.value

            Case ADODB.ParameterDirectionEnum.adParamReturnValue
                this.OuputParameters.Add param.value

        End Select
        
    Next

End Sub




