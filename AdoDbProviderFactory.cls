VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbProviderFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.Providers")

Option Explicit
'@PredeclaredId

Private Const ODBC_PROVIDER_NAME As String = "MSDASQL"
Private Const MICROSOFT_INDEXING_SERVICE_PROVIDER_NAME As String = "MSIDXS"
Private Const ACTIVE_DIRECTORY_PROVIDER_NAME As String = "ADSDSOObject"
Private Const MICROSOFT_JET_PROVIDER_NAME As String = "Microsoft.Jet.OLEDB.4.0"
Private Const MICROSOFT_ACE_PROVIDER_NAME As String = "Microsoft.ACE.OLEDB.12.0"
Private Const MICROSOFT_SQL_SERVER_PROVIDER_NAME As String = "SQLOLEDB"
Private Const ORACLE_PROVIDER_NAME As String = "MSDAORA"
Private Const INTERNET_PUBLISHING_PROVIDER_NAME As String = "MSDAIPP.DSO"
Private Const SIMPLE_DATASOURCE_PROVIDER_NAME As String = "MSDAOSP"

'See: https://docs.microsoft.com/en-us/sql/ado/guide/appendixes/appendix-a-providers?view=sql-server-2017
Public Enum AdoOleDbProviders
    OdbcProvider
    MicrosoftIndexingServiceProvider
    ActiveDirectoryServiceProvider
    MicrosoftJetProvider
    MicrosoftAceProvider
    MicrosoftSqlServerProvider
    OracleProvider
    InternetPublishingProvider
    SimpleDataSourceProvider
End Enum


Public Function GetProviderByName(ByVal providerName As String) As String

    Select Case providerName
    
        
        
        
    End Select
        
    
End Function


Public Function GetProviderByEnum(ByVal providerName As String)


End Function
