VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressView 
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6180
   OleObjectBlob   =   "ProgressView.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.ProgressBar")

Option Explicit

Public Event Cancel(ByVal CloseMode As VbQueryClose)

#If VBA7 Then  '64-Bit Windows and Either 64-Bit Or 32-Bit Excel
                        
    
    Private Declare PtrSafe Function GetWindowLong _
                           Lib "user32" Alias "GetWindowLongA" ( _
                           ByVal hwnd As LongPtr, _
                           ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong _
                           Lib "user32" Alias "SetWindowLongA" ( _
                           ByVal hwnd As LongPtr, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As LongPtr) As LongPtr

    Private Declare PtrSafe Function DrawMenuBar _
                           Lib "user32" ( _
                           ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindowA _
                           Lib "user32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As LongPtr


#Else '32-Bit Windows and 32-Bit Excel

    Private Declare Function GetWindowLong _
                            Lib "user32" Alias "GetWindowLongA" ( _
                            ByVal hwnd As Long, _
                            ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong _
                            Lib "user32" Alias "SetWindowLongA" ( _
                            ByVal hwnd As Long, _
                            ByVal nIndex As Long, _
                            ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar _
                            Lib "user32" ( _
                            ByVal hwnd As Long) As Long
    Private Declare Function FindWindowA _
                            Lib "user32" (ByVal lpClassName As String, _
                            ByVal lpWindowName As String) As Long

#End If



Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000

Private Sub CancelProcessBtn_Click()
    RaiseEvent Cancel(VBA.VbQueryClose.vbFormCode)
End Sub

Private Sub UserForm_Activate()
    HideTitleBar
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    RaiseEvent Cancel(CloseMode)
End Sub

Private Sub HideTitleBar()

#If VBA7 Then
    Dim lngWindow As LongPtr
    Dim lngFrmHdl As LongPtr
#Else
    Dim lngWindow As Long
    Dim lngFrmHdl As Long
#End If

    lngFrmHdl = FindWindowA(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lngFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lngFrmHdl, GWL_STYLE, lngWindow
    DrawMenuBar lngFrmHdl

End Sub
