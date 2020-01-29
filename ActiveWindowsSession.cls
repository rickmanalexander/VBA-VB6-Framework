VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveWindowsSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.ExtensibleLogging")

Option Explicit

Private Const CONNECTION_STRING As String = "ADsDSOObject"
Private Const CONNECTION_PROVIDER As String = "Active Directory Provider"
Private Const ADODB_OBJECT_STATE As Integer = 1     'the adStateOpen constant's value


Private Type TActiveWindowsSession
    userName As String
    UserDisplayName As String
    UserFirstName As String
    UserLastName As String
    UserCommonName As String
    UserEmailAddress As String
    UserTelephoneNumber As String
    UserDepartment As String
    CompanySiteName As String
    DomainName As String
    MachineName As String
    WindowsVerion As String
    AppVersion As String
End Type

Private this As TActiveWindowsSession

Private Sub Class_Initialize()
    GetUserAttributes
    GetSystemAttributes
End Sub

Private Sub Class_Terminate()
    With this
        .userName = vbNullString
        .UserDisplayName = vbNullString
        .UserFirstName = vbNullString
        .UserLastName = vbNullString
        .UserCommonName = vbNullString
        .UserEmailAddress = vbNullString
        .UserTelephoneNumber = vbNullString
        .UserDepartment = vbNullString
        .CompanySiteName = vbNullString
        .DomainName = vbNullString
        .MachineName = vbNullString
        .WindowsVerion = vbNullString
        .AppVersion = vbNullString
    End With
End Sub

Public Property Get userName() As String
    userName = this.userName
End Property

Public Property Get UserDisplayName() As String
    UserDisplayName = this.UserDisplayName
End Property

Public Property Get UserFirstName() As String
    UserFirstName = this.UserFirstName
End Property

Public Property Get UserLastName() As String
    UserLastName = this.UserLastName
End Property

Public Property Get UserCommonName() As String
    UserCommonName = this.UserCommonName
End Property

Public Property Get UserEmailAddress() As String
    UserEmailAddress = this.UserEmailAddress
End Property

Public Property Get UserDepartment() As String
    UserDepartment = this.UserDepartment
End Property

Public Property Get UserTelephoneNumber() As String
    UserTelephoneNumber = this.UserTelephoneNumber
End Property

Public Property Get CompanySiteName() As String
    CompanySiteName = this.CompanySiteName
End Property

Public Property Get DomainName() As String
    DomainName = this.DomainName
End Property

Public Property Get MachineName() As String
    MachineName = this.MachineName
End Property

Public Property Get WindowsVerion() As String
    WindowsVerion = this.WindowsVerion
End Property

Public Property Get AppVersion() As String
    AppVersion = this.AppVersion
End Property

Private Sub GetUserAttributes()

    Const scope As String = "subtree"

        this.userName = VBA.Environ$("userName")

        On Error GoTo CleanFail
        Dim rootDSE As Object
        Set rootDSE = GetObject("LDAP://RootDSE")
       
        Dim base As String
        base = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"

        'filter on user objects with the given account name
        Dim filter As String
        filter = "(&(objectClass=user)(objectCategory=Person)"

        filter = filter & "(sAMAccountName=" & this.userName & "))"
       
        Dim UserAttributes As String
        UserAttributes = "physicalDeliveryOfficeName,department,displayName," & _
                         "givenName,sn,mail,telephoneNumber"
       
        Dim conn As Object
        Set conn = CreateObject("ADODB.Connection")
            conn.provider = CONNECTION_STRING
            conn.Open CONNECTION_PROVIDER
       
        Dim cmnd As Object
        Set cmnd = CreateObject("ADODB.Command")

        Set cmnd.ActiveConnection = conn

            cmnd.commandText = base & ";" & filter & ";" _
                                & UserAttributes & ";" & scope
       
    Dim rs As Object
        Set rs = cmnd.Execute

        If rs.EOF Then GoTo CleanExit

        With this
            .CompanySiteName = GetFieldValueOrDefault(rs, "physicalDeliveryOfficeName")
            .UserDepartment = GetFieldValueOrDefault(rs, "department")
            .UserDisplayName = GetFieldValueOrDefault(rs, "displayName")
            .UserFirstName = GetFieldValueOrDefault(rs, "givenName")
            .UserLastName = GetFieldValueOrDefault(rs, "sn")
            .UserCommonName = Trim$(.UserFirstName & " " & .UserLastName)
            .UserEmailAddress = GetFieldValueOrDefault(rs, "mail")
            .UserTelephoneNumber = GetFieldValueOrDefault(rs, "telephoneNumber")
        End With

CleanExit:
    'bit-wise comparison
    If Not conn Is Nothing Then
        If (conn.State And ADODB_OBJECT_STATE) = ADODB_OBJECT_STATE Then
            conn.Close
        End If
    End If
    Set conn = Nothing

    If Not rs Is Nothing Then
        If (rs.State And ADODB_OBJECT_STATE) = ADODB_OBJECT_STATE Then
            rs.Close
        End If
    End If
    Set rs = Nothing
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Private Function GetFieldValueOrDefault(ByVal rst As Object, ByVal fieldName As String) As Variant
    On Error Resume Next
    GetFieldValueOrDefault = rst.Fields(fieldName).value
    On Error GoTo 0
End Function

Private Sub GetSystemAttributes()
    With this
        .DomainName = LCase$(Environ$("USERDNSDOMAIN"))
        .MachineName = Environ$("COMPUTERNAME")
        .WindowsVerion = Application.OperatingSystem
        .AppVersion = Application.Version
    End With
End Sub


Public Sub PrintToImmediateWindow()

    With this
        Debug.Print "Windows Verion: "; Tab(20); .WindowsVerion
        Debug.Print "App Version: "; Tab(20); .AppVersion
        Debug.Print "Machine Name: "; Tab(20); .MachineName
        Debug.Print "Site Name: "; Tab(20); .CompanySiteName
        Debug.Print "Domain DNS Name: "; Tab(20); .DomainName
        Debug.Print "User Name: "; Tab(20); .userName
        Debug.Print "Display Name: "; Tab(20); .UserDisplayName
        Debug.Print "First Name: "; Tab(20); .UserFirstName
        Debug.Print "Last Name: "; Tab(20); .UserLastName
        Debug.Print "Common Name: "; Tab(20); .UserCommonName
        Debug.Print "Email Address: "; Tab(20); .UserEmailAddress
    End With

End Sub

