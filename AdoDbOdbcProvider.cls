VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbOdbcProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.Providers")

Option Explicit
'@PredeclaredId

Private currentConnectionString As String

Public Function GetConnectionString(ByVal dataSourceName As String, _
                                    Optional ByVal hasHeaders As Boolean, _
                                    Optional ByVal treatAsText As Boolean) As String

    Dim ConnectionString As String
    ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dataSourceName
    
    Select Case dataSourceName
        Case IsExcelXlsFile(dataSourceName)
            ConnectionString = ConnectionString & ";Extended Properties='Excel 8.0"
            
            AppendExtendedProperties ConnectionString, hasHeaders, treatAsText
        
        Case IsExcelXlsxFile(dataSourceName)
            ConnectionString = ConnectionString & ";Extended Properties='Excel 12.0 Xml"
            
            AppendExtendedProperties ConnectionString, hasHeaders, treatAsText
            
        Case IsExcelMacroEnabledFile(dataSourceName)
            ConnectionString = ConnectionString & ";Extended Properties='Excel 12.0 Macro"
            
            AppendExtendedProperties ConnectionString, hasHeaders, treatAsText
        
        Case IsExcelBinaryFile(dataSourceName)
            ConnectionString = ConnectionString & ";Extended Properties='Excel 12.0"
            
            AppendExtendedProperties ConnectionString, hasHeaders, treatAsText
            
    End Select
      
End Function


Private Sub AppendExtendedProperties(ByRef ConnectionString As String, _
                                     ByVal hasHeaders As Boolean, _
                                     ByVal treatAsText As Boolean)
    
    If hasHeaders Then ConnectionString = ConnectionString & ";HDR=YES"
        
    If treatAsText Then ConnectionString = ConnectionString & ";IMEX=1"
    
    ConnectionString = ConnectionString & "'"

End Sub

Private Function IsExcelXlsFile(ByVal dataSourceName As String) As Boolean
    IsExcelXlsFile = (Right(Trim$(dataSourceName), 3) = "xls")
End Function

Private Function IsExcelXlsxFile(ByVal dataSourceName As String) As Boolean
    IsExcelXlsxFile = (Right(Trim$(dataSourceName), 4) = "xlsx")
End Function

Private Function IsExcelBinaryFile(ByVal dataSourceName As String) As Boolean
    IsExcelBinaryFile = (Right(Trim$(dataSourceName), 4) = "xlsb")
End Function

Private Function IsExcelMacroEnabledFile(ByVal dataSourceName As String) As Boolean
    IsExcelMacroEnabledFile = (Right(Trim$(dataSourceName), 4) = "xlsm")
End Function
