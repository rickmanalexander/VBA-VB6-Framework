VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework.ProgressBar")

Option Explicit

Private WithEvents View As ProgressView
Attribute View.VB_VarHelpID = -1

#If VBA7 Then  '64-Bit Windows and Either 64-Bit Or 32-Bit Excel
                     
    'Note: https://codekabinett.com/rdumps.php?Lang=2&targetDoc=windows-api-declaration-vba-64-bit
    #If Win64 Then
        Private Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As LongPtr
    #Else
        Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
    #End If


#Else '32-Bit Windows and 32-Bit Excel
              
    Private Declare Function GetTickCount _
                            Lib "kernel32.dll" () As Long

#End If

Private Const PROGINDICATOR_MAXHEIGHT As Integer = 142.75
Private Const PARENTPROCSTATUS_MAXHEIGHT As Integer = 10
Private Const PROCSTATUS_MAXTOP As Integer = 16
Private Const PROGRESSBAR_MAXTOP As Integer = 41
Private Const PROGRESSBAR_MAXWIDTH As Integer = 270
Private Const ELAPSEDTIME_MAXTOP As Integer = 83
Private Const TIMEREMAINING_MAXTOP As Integer = 94

Private Const STARTPOS_LEFT_OFFSET As Single = 0.5
Private Const STARTPOS_TOP_OFFSET As Single = 0.5

Private Const ERR_ORPHANPROC_NOPARENT As String = "You specified that this proccess has a parent, " & _
                                                  "but you are using the 'UpdateOrphanProgress' method"
Private Const ERR_HASPARENT_NOTSPECIFIED As String = "You specified that this proccess does not have a parent, " & _
                                                     "but you are using the 'UpdateParentChildProgress' method."
Private Const ERR_INVALIDPROGPERCENT As String = "Either the CurrProgCnt equals 0, is greater than 0, or it " & _
                                                  "is greater than TotalProgCnt."
Private Const ERR_INVALIDPARENTCOUNT As String = "Either the ParentCurrCnt equals 0, is greater than 0, or it " & _
                                                  "is greater than ParentTotalCnt."

Private Enum ProgressIndicatorError    'should be Private ??
    Error_OrphanProcHasParent = vbObjectError + 1001
    Error_HasParentProcNotSpecified
    Error_InvalidProgressPercentage
    Error_InvalidParentCount
End Enum

Private Type TProgressIndicator
    StartTime As Double
    TimeElapsed As Double
    SecondsElapsed As Double
    MinutesElapsed As Double
    HoursElapsed As Double
    SecondsRemaining As Double
    MinutesRemaining As Double
    HoursRemaining As Double
    ItemsRemaining As Double
    ParentChildIterationCount As Long
    HasParentProccess As Boolean
    CanCancel As Boolean
    Cancelling As Boolean
    CalculateExecutionTime As Boolean
    PercentComplete As Double
End Type

Private this As TProgressIndicator

Public Property Get ShouldCancel() As Boolean
    If IsCancelRequested Then ShouldCancel = True
End Property

Private Property Get IsCancelRequested() As Boolean
    IsCancelRequested = this.Cancelling
End Property

Private Sub Class_Terminate()
    If Not View Is Nothing Then: If View.Visible Then View.Hide
End Sub

Public Sub LoadProgIndicator(Optional ByVal HasParentProccess As Boolean, _
                              Optional ByVal CanCancel As Boolean, _
                              Optional ByVal CalculateExecutionTime As Boolean)
    
    this.CalculateExecutionTime = CalculateExecutionTime
    
    If this.CalculateExecutionTime Then this.StartTime = GetTickCount()
    
    this.HasParentProccess = HasParentProccess
    
    this.CanCancel = CanCancel
    
    Set View = New ProgressView
    
    With View
        If this.HasParentProccess Then
            .height = PROGINDICATOR_MAXHEIGHT
            .ParentProcedureStatus.height = PARENTPROCSTATUS_MAXHEIGHT
            .ProcedureStatus.Top = PROCSTATUS_MAXTOP
            .frameProgressBar.Top = PROGRESSBAR_MAXTOP
            
            .lblElapsedTime.Top = ELAPSEDTIME_MAXTOP
            .ElapsedTime.Top = ELAPSEDTIME_MAXTOP
            
            .lblTimeRemaining.Top = TIMEREMAINING_MAXTOP
            .TimeRemaining.Top = TIMEREMAINING_MAXTOP
        End If
        If this.CanCancel Then .CancelProcessBtn.Visible = True: .CancelProcessBtn.Enabled = True
        .ProgressBar.width = 0
        .StartUpPosition = 0
        .Left = Application.Left + (STARTPOS_LEFT_OFFSET * Application.width) - (STARTPOS_LEFT_OFFSET * .width)
        .Top = Application.Top + (STARTPOS_TOP_OFFSET * Application.height) - (STARTPOS_TOP_OFFSET * .height)
        .Show
    End With
        
End Sub


Public Sub UpdateOrphanProgress(ByRef ProgStatusText As Variant, _
                                 ByRef CurrProgCnt As Long, _
                                 ByRef TotalProgCnt As Long)
    
        ThrowIfOrphanProcHasParent
        
        ThrowIfInvalidProgPercent CurrProgCnt, TotalProgCnt
        
        this.PercentComplete = CurrProgCnt / TotalProgCnt

            With View
                .ProcedureStatus.Caption = ProgStatusText & " " & _
                                           CurrProgCnt & " of " & TotalProgCnt
                
                .ProgressBar.width = this.PercentComplete * PROGRESSBAR_MAXWIDTH
                
            End With
        
        If this.CalculateExecutionTime Then CalculateTime CurrProgCnt, TotalProgCnt

        DoEvents
        
        If CurrProgCnt = TotalProgCnt Then View.Hide

End Sub


Public Sub UpdateParentChildProgress(ByRef ParentProgStatusText As Variant, _
                                      ByRef ParentCurrCnt As Long, _
                                      ByRef ParentTotalCnt As Long, _
                                      ByRef ChildProgStatusText As Variant, _
                                      ByRef ChildCurrProgCnt As Long, _
                                      ByRef ChildProgCnt As Long, _
                                      ByRef TotalProgCnt As Long)

        ThrowIfHasParentNotSpecified
        
        ThrowIfInvalidParentCount ParentCurrCnt, ParentTotalCnt
        
        ThrowIfInvalidProgPercent ChildCurrProgCnt, ChildProgCnt
        
        this.ParentChildIterationCount = this.ParentChildIterationCount + 1
        
        this.PercentComplete = ChildCurrProgCnt / ChildProgCnt

            With View
                .ParentProcedureStatus.Caption = ParentProgStatusText & " " & _
                                                 ParentCurrCnt & " of " & ParentTotalCnt
                                                     
                .ProcedureStatus.Caption = ChildProgStatusText & " " & _
                                           ChildCurrProgCnt & " of " & ChildProgCnt

                .ProgressBar.width = this.PercentComplete * PROGRESSBAR_MAXWIDTH
            End With
        
        If this.CalculateExecutionTime Then CalculateTime this.ParentChildIterationCount, TotalProgCnt
        
        DoEvents
        
        If ParentCurrCnt = ParentTotalCnt Then
           If ChildCurrProgCnt = ChildProgCnt Then View.Hide
        End If

End Sub

Private Sub CalculateTime(ByRef CurrProgCntIn As Long, ByRef TotalProgCntIn As Long)
        
    With this
        If CurrProgCntIn = TotalProgCntIn Then
    
            View.ElapsedTime.Caption = .HoursElapsed & " hours, " & _
                                       .MinutesElapsed & " minutes, " & .SecondsElapsed & " seconds"
                                     
            View.TimeRemaining.Caption = "0 hours, " & 0 & _
                                         " minutes, " & 0 & " seconds"
        Else
            .TimeElapsed = (GetTickCount() - this.StartTime)
            
            .SecondsElapsed = .TimeElapsed / 1000
            .MinutesElapsed = RoundTime(.TimeElapsed, 60000)
            .HoursElapsed = RoundTime(.TimeElapsed, 3600000)
         
            .ItemsRemaining = TotalProgCntIn - CurrProgCntIn
    
            .SecondsRemaining = (.SecondsElapsed * (TotalProgCntIn / CurrProgCntIn)) - .SecondsElapsed
            .MinutesElapsed = RoundTime(.SecondsRemaining, 60)
            .HoursElapsed = RoundTime(.SecondsRemaining, 60)
            
            View.ElapsedTime.Caption = .HoursElapsed & " hours, " & _
                                       .MinutesElapsed & " minutes, " & .SecondsElapsed & " seconds"
                                     
            View.TimeRemaining.Caption = .HoursRemaining & " hours, " & .MinutesRemaining & _
                                         " minutes, " & .SecondsRemaining & " seconds"
        End If
    End With
    
End Sub
    
Private Function RoundTime(ByRef TimeElapsedIn As Double, ByVal IntervalIn As Long) As Double
    RoundTime = Int(TimeElapsedIn / IntervalIn)
End Function

Private Sub ThrowIfOrphanProcHasParent()
    If this.HasParentProccess Then
        Beep
        Err.Raise ProgressIndicatorError.Error_OrphanProcHasParent, _
                    VBA.TypeName(Me), ERR_ORPHANPROC_NOPARENT
    End If
End Sub

Private Sub ThrowIfHasParentNotSpecified()
    If Not this.HasParentProccess Then
        Beep
        Err.Raise ProgressIndicatorError.Error_HasParentProcNotSpecified, _
                  VBA.TypeName(Me), ERR_HASPARENT_NOTSPECIFIED
    End If
End Sub
        

Private Sub ThrowIfInvalidProgPercent(ByRef CurrProgCntIn As Long, ByRef TotalProgCntIn As Long)
    If Not (CurrProgCntIn > 0 And CurrProgCntIn <= TotalProgCntIn) Then
        Beep
        Err.Raise ProgressIndicatorError.Error_InvalidProgressPercentage, _
                  VBA.TypeName(Me), ERR_INVALIDPROGPERCENT
    End If
End Sub


Private Sub ThrowIfInvalidParentCount(ByRef ParentCurrCntIn As Long, ByRef ParentTotalCntIn As Long)
    If Not (ParentCurrCntIn > 0 And ParentCurrCntIn <= ParentTotalCntIn) Then
        Beep
        Err.Raise ProgressIndicatorError.Error_InvalidParentCount, _
                  VBA.TypeName(Me), ERR_INVALIDPARENTCOUNT
    End If
End Sub

Private Sub View_Cancel(ByVal CloseMode As VbQueryClose)
    If CloseMode = VBA.VbQueryClose.vbFormControlMenu Or _
       CloseMode = VBA.VbQueryClose.vbFormCode Then
        If this.CanCancel Then
            If MsgBox("Are you sure that you want to Cancel this operation?", _
                      vbQuestion + vbYesNo, "Process Dialog") = vbYes Then
                View.Hide
                this.Cancelling = True
            End If
        End If
    End If
End Sub

'Private Sub View_Cancel(ByVal CloseMode As VbQueryClose)
'
'    Select Case CloseMode
'
'        Case VBA.VbQueryClose.vbFormControlMenu
'            If this.CanCancel Then
'                If MsgBox("Are you sure that you want to Cancel this operation?", _
'                          vbQuestion + vbYesNo, "Process Dialog") = vbYes Then
'                    View.Hide
'                    this.Cancelling = True
'                End If
'            End If
'
'        Case VBA.VbQueryClose.vbFormCode
'            If this.CanCancel Then
'                If MsgBox("Are you sure that you want to Cancel this operation?", _
'                          vbQuestion + vbYesNo, "Process Dialog") = vbYes Then
'                    View.Hide
'                    this.Cancelling = True
'                End If
'            End If
'
'    End Select
'
'End Sub



