VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.Extensions")
Option Explicit

Public Enum UnitOfWorkErrors
    DidNotInitializeViaInit = vbObjectError + 1024
    TransactionsNotSupported
End Enum

Private WithEvents CurrentConnection As ADODB.connection
Attribute CurrentConnection.VB_VarHelpID = -1

Private Type TUnitOfWork
    closeConnectionAfterCommit As Boolean
    calledInit As Boolean
    CanConnect As Boolean
    SuccessfulConnection As Boolean
    SuccessfulDisconnection As Boolean
    TransactionHasBegun As Boolean
    TranactionComplete As Boolean
    TransactionRolledBack  As Boolean
    CommandWillExecute As Boolean
    CommandComplete As Boolean
End Type

Private this As TUnitOfWork

Public Sub Init(ByRef connection As ADODB.connection, _
                Optional ByVal closeConnectionOnCommit As Boolean = False)
    
    ThrowIfNoTransactionSupport connection
        
    this.calledInit = True
    
    Set CurrentConnection = connection
    
    this.closeConnectionAfterCommit = closeConnectionOnCommit
    
End Sub


Public Property Get ActiveConnection() As ADODB.connection
    Set ActiveConnection = CurrentConnection
End Property

Public Property Get CanConnect() As Boolean
    CanConnect = this.CanConnect
End Property

Public Property Get SuccessfulConnection() As Boolean
    SuccessfulConnection = this.SuccessfulConnection
End Property

Public Property Get SuccessfulDisconnection() As Boolean
    SuccessfulDisconnection = this.SuccessfulDisconnection
End Property

Public Property Get TransactionHasBegun() As Boolean
    TransactionHasBegun = this.TransactionHasBegun
End Property

Public Property Get TranactionComplete() As Boolean
    TranactionComplete = this.TranactionComplete
End Property

Public Property Get TransactionRolledBack() As Boolean
    TransactionRolledBack = this.TransactionRolledBack
End Property

Public Property Get CommandWillExecute() As Boolean
    CommandWillExecute = this.CommandWillExecute
End Property

Public Property Get CommandComplete() As Boolean
    CommandComplete = this.CommandComplete
End Property


Public Sub Begin()
    
    ThrowIfInitNotCalled

    On Error GoTo CleanFail
    CurrentConnection.BeginTrans
    
CleanExit:
    Exit Sub
    
CleanFail:
    TryRollBack     'will this cause an infinte loop??
    Resume CleanExit
    
End Sub


Public Sub Commit()

    ThrowIfInitNotCalled
    
    On Error GoTo CleanFail
    CurrentConnection.CommitTrans
    
    If this.closeConnectionAfterCommit Then
        If Not CurrentConnection Is Nothing Then
            If (CurrentConnection.State And adStateOpen) = adStateOpen Then CurrentConnection.Close
            Set CurrentConnection = Nothing
            
        End If
        
    End If
    
CleanExit:
    Exit Sub
    
CleanFail:
    TryRollBack         'will this cause an infinte loop??
    Resume CleanExit
    
End Sub


Public Sub Rollback()

    ThrowIfInitNotCalled
    
    On Error GoTo CleanFail
    CurrentConnection.RollbackTrans
    
CleanExit:
    Exit Sub
    
CleanFail:
    Resume CleanExit
    
End Sub


Private Sub TryRollBack()

    On Error Resume Next
    CurrentConnection.RollbackTrans

End Sub


Private Sub ThrowIfNoTransactionSupport(ByRef connection As ADODB.connection)

    Dim dbProvider As String
    dbProvider = connection.provider

    If Not SupportsTransactions(connection) Then
            Err.Raise Number:=UnitOfWorkErrors.TransactionsNotSupported, _
                      Source:=TypeName(Me), _
                      Description:="Database Provider: " & dbProvider & " does Not support use of transactions"
    End If

End Sub



Private Sub ThrowIfInitNotCalled()

    If Not this.calledInit Then
        Err.Raise Number:=UnitOfWorkErrors.DidNotInitializeViaInit, _
                  Source:=TypeName(Me), _
                  Description:="You must call the " & TypeName(Me) & " '.Init method to use this class"
    End If

End Sub

'Connection Status Events
Private Sub CurrentConnection_WillConnect(ConnectionString As String, _
                                          UserID As String, Password As String, _
                                          Options As Long, adStatus As ADODB.EventStatusEnum, _
                                          ByVal pConnection As ADODB.connection)
                                           
    this.CanConnect = (adStatus = EventStatusEnum.adStatusOK)

End Sub

Private Sub CurrentConnection_ConnectComplete(ByVal pError As ADODB.Error, _
                                       adStatus As ADODB.EventStatusEnum, _
                                       ByVal pConnection As ADODB.connection)
                                       
    this.SuccessfulConnection = (adStatus = EventStatusEnum.adStatusOK)

End Sub

Private Sub CurrentConnection_Disconnect(adStatus As ADODB.EventStatusEnum, _
                                  ByVal pConnection As ADODB.connection)

    this.SuccessfulDisconnection = (adStatus = EventStatusEnum.adStatusOK)

    
End Sub


'Transaction Events
Private Sub CurrentConnection_BeginTransComplete(ByVal TransactionLevel As Long, _
                                                 ByVal pError As ADODB.Error, _
                                                 adStatus As ADODB.EventStatusEnum, _
                                                 ByVal pConnection As ADODB.connection)
                                          
    this.TransactionHasBegun = (adStatus = EventStatusEnum.adStatusOK)

End Sub

Private Sub CurrentConnection_CommitTransComplete(ByVal pError As ADODB.Error, _
                                                  adStatus As ADODB.EventStatusEnum, _
                                                  ByVal pConnection As ADODB.connection)
                                           
    this.TranactionComplete = (adStatus = EventStatusEnum.adStatusOK)

    
End Sub

Private Sub CurrentConnection_RollbackTransComplete(ByVal pError As ADODB.Error, _
                                                    adStatus As ADODB.EventStatusEnum, _
                                                    ByVal pConnection As ADODB.connection)
                                             
    this.TransactionRolledBack = (adStatus = EventStatusEnum.adStatusOK)

End Sub


'Command Events
Private Sub CurrentConnection_WillExecute(Source As String, _
                                          cursorType As ADODB.CursorTypeEnum, _
                                          lockType As ADODB.LockTypeEnum, _
                                          Options As Long, adStatus As ADODB.EventStatusEnum, _
                                          ByVal pCommand As ADODB.Command, _
                                          ByVal pRecordset As ADODB.recordset, _
                                          ByVal pConnection As ADODB.connection)
                                   
    this.CommandWillExecute = (adStatus = EventStatusEnum.adStatusOK)

End Sub

Private Sub CurrentConnection_ExecuteComplete(ByVal RecordsAffected As Long, _
                                              ByVal pError As ADODB.Error, _
                                              adStatus As ADODB.EventStatusEnum, _
                                              ByVal pCommand As ADODB.Command, _
                                              ByVal pRecordset As ADODB.recordset, _
                                              ByVal pConnection As ADODB.connection)
                                       
    this.CommandComplete = (adStatus = EventStatusEnum.adStatusOK)
    
End Sub


Private Sub CurrentConnection_InfoMessage(ByVal pError As ADODB.Error, _
                                          adStatus As ADODB.EventStatusEnum, _
                                          ByVal pConnection As ADODB.connection)
                                     

End Sub



