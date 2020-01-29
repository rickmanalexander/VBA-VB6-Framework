VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetworkConnectionMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.ExtensibleLogging")

Option Explicit

Private Const NULL_VALUE As String = "null"

Private Const WIFI_CONNECTION_ID As String = "Wi-Fi"
Private Const ETHERNET_V1_CONNECTION_ID As String = "Ethernet"
Private Const ETHERNET_V2_CONNECTION_ID As String = "Ethernet 2"
Private Const BLUETOOTH_CONNECTION_ID As String = "Bluetooth Network Connection"
Private Const LAN_CONNECTION_ID As String = "Local Area Connection"

Private Type TNetworkConnectionMonitor
    HasWiFi As Boolean
    HasEthernetV1 As Boolean
    HasEthernetV2 As Boolean
    HasBluetooth As Boolean
    HasLocalAreaConnection As Boolean
    ConnectedToInternet As Boolean
End Type

Private Enum NetworkConnectionStatusEnum
    Disconnected
    Connecting
    Connected
    Disconnecting
    HardwareNotPresent
    HardwareDisabled
    HardwareMalfunction
    MediaDisconnected
    Authenticating
    AuthenticationSucceeded
    AuthenticationFailed
    InvalidAddress
    CredentialsRequired
End Enum

Private this As TNetworkConnectionMonitor

Public Property Get HasWiFi()
    HasWiFi = this.HasWiFi
End Property

Public Property Get HasEthernetV1()
    HasEthernetV1 = this.HasEthernetV1
End Property

Public Property Get HasEthernetV2()
    HasEthernetV2 = this.HasEthernetV2
End Property

Public Property Get HasBluetooth()
    HasBluetooth = this.HasBluetooth
End Property

Public Property Get HasLocalAreaConnection()
    HasLocalAreaConnection = this.HasLocalAreaConnection
End Property

Public Property Get ConnectedToInternet()
    ConnectedToInternet = this.ConnectedToInternet
End Property

Private Sub Class_Initialize()
    DetermineConnectionStatus
End Sub

Private Sub DetermineConnectionStatus()

    Dim WinMgmtInsrumentation As Object, WinNtwkAdapter As Object, Instance As Object
    
    Set WinMgmtInsrumentation = GetObject("WINMGMTS:\\.\ROOT\cimv2")
    
    Set WinNtwkAdapter = WinMgmtInsrumentation.InstancesOf("Win32_NetworkAdapter")
    
    On Error GoTo CleanFail
    For Each Instance In WinNtwkAdapter
        If Instance.NetconnectionID <> NULL_VALUE Then
            
            Select Case Instance.NetconnectionID
                'conduct bit wise comparisons
                Case WIFI_CONNECTION_ID
                    If (Instance.NetConnectionStatus And _
                        NetworkConnectionStatusEnum.Connected) Then this.HasWiFi = True
                
                Case ETHERNET_V1_CONNECTION_ID
                    If (Instance.NetConnectionStatus And _
                        NetworkConnectionStatusEnum.Connected) Then this.HasEthernetV1 = True
                
                Case ETHERNET_V2_CONNECTION_ID
                    If (Instance.NetConnectionStatus And _
                        NetworkConnectionStatusEnum.Connected) Then this.HasEthernetV2 = True
                
                Case BLUETOOTH_CONNECTION_ID
                    If (Instance.NetConnectionStatus And _
                        NetworkConnectionStatusEnum.Connected) Then this.HasBluetooth = True
                
                Case LAN_CONNECTION_ID
                    If (Instance.NetConnectionStatus And _
                        NetworkConnectionStatusEnum.Connected) Then this.HasLocalAreaConnection = True
                    
            End Select
        
        End If
    Next
    
    'If None of these are true then not connected to the internet
    this.ConnectedToInternet = (this.HasEthernetV2 Or this.HasWiFi Or this.HasEthernetV1)

CleanExit:
    Exit Sub
    
CleanFail:
    Resume CleanExit
    
End Sub


Public Function IsConnectedToHost(ByVal HostName As String, ByVal PingCount As Integer, _
                                  Optional ByVal PingTimeOut As Long = 550)

    With CreateObject("WScript.Shell")
        IsConnectedToHost = (.Run("%comspec% /c ping.exe -n " & PingCount & _
                                 " -w " & PingTimeOut & " " & HostName & _
                                 " | find ""TTL="" > nul 2>&1", 0, True) = 0)
    End With

End Function
