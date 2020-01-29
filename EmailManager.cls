VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.ExtensibleLogging")

Option Explicit

Public Function IsOutlookAvailiable() As Boolean
    Dim OutLookApp As Object
      On Error Resume Next
      Set OutLookApp = CreateObject("Outlook.Application")
      IsOutlookAvailiable = Not OutLookApp Is Nothing
      On Error GoTo 0
End Function

Public Sub SendOutlookEmail(ByVal EmailToAddress As String, ByVal EmailSubject As String, _
                            ByVal EmailBody As String)

    Dim OutLookApp As Object, OutMail As Object
        
        On Error GoTo CleanFail
        Set OutLookApp = CreateObject("Outlook.Application")
        
        Set OutMail = OutLookApp.CreateItem(0)

            On Error GoTo CleanFail
            With OutMail
                .To = EmailToAddress
                .Subject = EmailSubject
                .Body = EmailBody
                .Send
            End With

CleanExit:
    Set OutLookApp = Nothing
    Set OutMail = Nothing
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Sub SendCDOEmail(ByVal SMTPServerAddress As String, ByVal FromAddress As String, _
                        ByVal ToAddress As String, ByVal Subject As String, ByVal Body As String, _
                        Optional SMTPServerPort As Integer = 25)

    Dim CDOMsg As Object, CDOConfig As Object, ConfigFlds As Object

    On Error GoTo CleanFail
    Set CDOMsg = CreateObject("CDO.Message")
    Set CDOConfig = CreateObject("CDO.Configuration")

    CDOConfig.Load -1     ' CDO Source Defaults
    Set ConfigFlds = CDOConfig.Fields
        With ConfigFlds
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServerAddress
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort
            .Update
        End With

        With CDOMsg
            Set .Configuration = CDOConfig
            .To = ToAddress
            .from = Chr(34) & VBA.Environ$("USERNAME") & Chr(34) & " <" & _
                              FromAddress & ">"
            .Subject = Subject

            'Set importance or Priority to high
            .Fields("urn:schemas:httpmail:importance") = 2
            .Fields("urn:schemas:mailheader:X-Priority") = 1
            .Fields.Update

            .TextBody = Body

            .Send
        End With

CleanExit:
    Set CDOMsg = Nothing
    Set CDOConfig = Nothing
    Set ConfigFlds = Nothing
    Exit Sub

CleanFail:
    Resume CleanExit
    
End Sub

