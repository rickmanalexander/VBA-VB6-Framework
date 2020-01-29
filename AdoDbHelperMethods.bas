Attribute VB_Name = "AdoDbHelperMethods"
'@Folder("Framework.DataAccess.Common.Extensions")
Option Explicit

Public Function SupportsTransactions(ByRef connection As ADODB.connection) As Boolean
    
    Const TRANSACTION_PROPERTY_NAME As String = "Transaction DDL"
    SupportsTransactions = ConnectionPropertyExists(connection, TRANSACTION_PROPERTY_NAME)
    
End Function


Public Function ConnectionPropertyExists(ByRef connection As ADODB.connection, ByVal propertyName As String) As Boolean

    On Error Resume Next
    ConnectionPropertyExists = connection.Properties(propertyName)
    Err.Clear
    
End Function

Public Function RecordsetPropertyExists(ByRef recordset As ADODB.recordset, ByVal propertyName As String) As Boolean

    On Error Resume Next
    RecordsetPropertyExists = recordset.Properties(propertyName)
    Err.Clear

End Function
