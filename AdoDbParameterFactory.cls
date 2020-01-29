VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDbParameterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.AdoDbCommand.Parameters")

Option Explicit
'@PredeclaredId

Public Enum AdoDbParameterFactoryErrors
    NonDefaultInstanceCreated = vbObjectError + 1024
End Enum

Public Function Create(ByVal name As String, _
                       ByVal dataType As ADODB.DataTypeEnum, _
                       ByRef value As Variant, _
                       ByVal direction As ADODB.ParameterDirectionEnum, _
                       Optional ByVal size As Long, _
                       Optional ByVal precision As Byte, _
                       Optional ByVal numericScale As Byte) As ADODB.parameter
            
            
    Set Create = New ADODB.parameter
    With Create
        .name = name
        .Type = dataType
        .value = value
        .direction = direction
        .size = size
        .precision = precision
        .numericScale = numericScale
    End With
            
End Function


Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is AdoDbParameterFactory
End Property


Private Sub Class_Initialize()
    If Not IsDefaultInstance Then Err.Raise Number:=AdoDbParameterFactoryErrors.NonDefaultInstanceCreated, _
                                            Source:=TypeName(Me), _
                                            Description:="Non-default instances of a factory class are invalid"
End Sub
