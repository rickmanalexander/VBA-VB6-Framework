VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelApplicationProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framework")
Option Explicit

'These constants define the default restoration behaviors for the class
Private Const DEFAULT_RESTORE_CALCULATION As Boolean = True
Private Const DEFAULT_RESTORE_DISPLAY_ALERTS As Boolean = True
Private Const DEFAULT_RESTORE_ENABLE_EVENTS As Boolean = True
Private Const DEFAULT_RESTORE_SCREEN_UPDATING As Boolean = True

'Set this to true to ensure a persisted state is restored, _
 even if the client code forgets to restore
Private Const DEFAULT_RESTORE_ON_TERMINATE As Boolean = True

Public Enum ExcelPropertiesError
    Error_AlreadyInitialized = vbObjectError + 42
    Error_AppPropertiesNotSaved
End Enum

Private Type TExcelApplicationProperties
    Calculation As XlCalculationState
    DisplayAlerts As Boolean
    EnableEvents As Boolean
    ScreenUpdating As Boolean
    RestoreCalculation As Boolean
    RestoreDisplayAlerts As Boolean
    RestoreEnableEvents As Boolean
    RestoreScreenUpdating As Boolean
    RestoreOnTerminate As Boolean
    Saved As Boolean
    IsRestored As Boolean
End Type

Private this As TExcelApplicationProperties

'Set the default restoration behaviours on intialize
Private Sub Class_Initialize()
    With this
        .RestoreCalculation = DEFAULT_RESTORE_CALCULATION
        .RestoreDisplayAlerts = DEFAULT_RESTORE_DISPLAY_ALERTS
        .RestoreEnableEvents = DEFAULT_RESTORE_ENABLE_EVENTS
        .RestoreScreenUpdating = DEFAULT_RESTORE_SCREEN_UPDATING
    End With
End Sub

'By default, restore the settings if we didn't do it explicitly
Private Sub Class_Terminate()
    If this.Saved And Not this.IsRestored And DEFAULT_RESTORE_ON_TERMINATE Then
        Me.Restore
    End If
End Sub

Public Property Get RestoreCalculation() As Boolean
  RestoreCalculation = this.RestoreCalculation
End Property

Public Property Let RestoreCalculation(ByVal value As Boolean)
  this.RestoreCalculation = value
End Property


Public Property Get RestoreDisplayAlerts() As Boolean
  RestoreDisplayAlerts = this.RestoreDisplayAlerts
End Property

Public Property Let RestoreDisplayAlerts(ByVal value As Boolean)
  this.RestoreDisplayAlerts = value
End Property


Public Property Get RestoreEnableEvents() As Boolean
  RestoreEnableEvents = this.RestoreEnableEvents
End Property

Public Property Let RestoreEnableEvents(ByVal value As Boolean)
  this.RestoreEnableEvents = value
End Property


Public Property Get RestoreScreenUpdating() As Boolean
  RestoreScreenUpdating = this.RestoreScreenUpdating
End Property

Public Property Let RestoreScreenUpdating(ByVal value As Boolean)
  this.RestoreScreenUpdating = value
End Property


Public Property Get RestoreOnTerminate() As Boolean
  RestoreOnTerminate = this.RestoreOnTerminate
End Property

Public Property Let RestoreOnTerminate(ByVal value As Boolean)
  this.RestoreOnTerminate = value
End Property

Public Sub Save()
  If Not this.Saved Then
    'Save all of the settings
    With Application
      this.Calculation = .Calculation
      this.DisplayAlerts = .DisplayAlerts
      this.EnableEvents = .EnableEvents
      this.ScreenUpdating = .ScreenUpdating
    End With
    this.Saved = True
  Else
    Err.Raise ExcelPropertiesError.Error_AlreadyInitialized, VBA.TypeName(Me), "Properties have already been saved."
  End If
End Sub

Public Sub Restore()
  'Only restore the settings that we want restored
  '(which by default is all of them)
  With Application
    If this.RestoreCalculation Then .Calculation = this.Calculation
    If this.RestoreDisplayAlerts Then .DisplayAlerts = this.DisplayAlerts
    If this.RestoreEnableEvents Then .EnableEvents = this.EnableEvents
    If this.RestoreScreenUpdating Then .ScreenUpdating = this.ScreenUpdating
  End With
  this.IsRestored = True
End Sub

Public Sub ToggleProperties()

    If this.Saved Then
        With Application
            If .Calculation = xlCalculationAutomatic Then .Calculation = xlCalculationManual
            If .DisplayAlerts Then .DisplayAlerts = False
            If .EnableEvents Then .EnableEvents = False
            If .ScreenUpdating Then .ScreenUpdating = False
        End With
    Else
        Err.Raise ExcelPropertiesError.Error_AppPropertiesNotSaved, _
                  VBA.TypeName(Me), "Initial Application Properties were not saved."
    End If
    
End Sub
