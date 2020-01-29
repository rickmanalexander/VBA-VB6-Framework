VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAdoDbParameterMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess")
Option Explicit

Public Sub SetParameters(ByRef commandReturn As ADODB.Command, ByRef parameterValues As Variant)
End Sub

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1024, Source:=TypeName(Me), Description:="An Interface class must not be instantiated."
End Sub

