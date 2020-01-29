VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitOfWorkFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.DataAccess.Common.Extensions")

Option Explicit
'@PredeclaredId

Public Enum UnitOfWorkFactoryErrors
    NonDefaultInstanceCreated = vbObjectError + 1024
End Enum

Public Function Create(connection As ADODB.connection, _
                       Optional ByVal closeConnectionOnCommit As Boolean = False) As UnitOfWork
                       
    Set Create = New UnitOfWork
    Create.Init connection, closeConnectionOnCommit
    
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is UnitOfWorkFactory
End Property


Private Sub Class_Initialize()
    If Not IsDefaultInstance Then Err.Raise Number:=UnitOfWorkFactoryErrors.NonDefaultInstanceCreated, _
                                            Source:=TypeName(Me), _
                                            Description:="Non-default instances of a factory class are invalid"
End Sub

