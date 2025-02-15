VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Event MouseClickOutside()
Event AppDeactivate()
Event CheckForAdditionalCloseConditions()
Event InitialMouseUp(ByVal Button As Integer)

Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents tmrPopUp As cTimer
Attribute tmrPopUp.VB_VarHelpID = -1

Private mMouseDown As Boolean, mInitialMouseDownKey As Long

Public Sub Load(Widget As Object, ByVal ScreenX As Long, ByVal ScreenY As Long, ByVal WidthPxl As Long, ByVal HeightPxl As Long, Optional ByVal Zoom As Single = 1, Optional ByVal WithDropShadow As Boolean = True)
  Unload
  If Widget Is Nothing Then Exit Sub

  If TypeOf Widget Is cwMenu Then
    If Not TypeOf Widget.InitiatorWidget.object Is cwMenu Then
      If Not fActivePopUp Is Nothing Then If Not fActivePopUp Is Me Then fActivePopUp.Unload
      Set fActivePopUp = Me
    End If
  Else
    If Not fActivePopUp Is Nothing Then If Not fActivePopUp Is Me Then fActivePopUp.Unload
    Set fActivePopUp = Me
  End If
  
  If Widget.Widget.Root Is Nothing Then
    Set Form = Cairo.WidgetForms.CreateChild(0, True, False, False, WithDropShadow)
  Else
    Set Form = Cairo.WidgetForms.CreateChild(Widget.Widget.Root.DialogFormHwnd, True, False, False, WithDropShadow)
  End If
  Form.WidgetRoot.Zoom = Zoom
  Form.Widgets.Add Widget, "CurPopupWidget", 0, 0, WidthPxl / Zoom, HeightPxl / Zoom
  Form.Move ScreenX, ScreenY, WidthPxl, HeightPxl
  mInitialMouseDownKey = Form.WidgetRoot.MouseKeyDown
  mMouseDown = True

  Set tmrPopUp = New_c.Timer(20, True)
End Sub

Public Sub Show()
  If Not Form Is Nothing Then Form.Show
End Sub

Public Sub Unload()
  On Error Resume Next
    Set tmrPopUp = Nothing
    If fActivePopUp Is Me Then Set fActivePopUp = Nothing
  
    If Not Form Is Nothing Then
      If Not Form.Widgets Is Nothing Then Form.Widgets.RemoveAll
      If Form.hWnd Then Form.Unload
    End If
    Set Form = Nothing
  If Err Then Err.Clear
End Sub

Private Sub Form_ActivateApp(ByVal Activated As Boolean, ByVal OtherThreadID As Long)
  If Not Activated Then
    On Error Resume Next
      RaiseEvent AppDeactivate
    If Err Then Err.Clear
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  mMouseDown = True
End Sub
 
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If mMouseDown And mInitialMouseDownKey > 0 Then
    RaiseEvent InitialMouseUp(mInitialMouseDownKey)
    mInitialMouseDownKey = 0
  End If
  mMouseDown = False
End Sub

Private Sub tmrPopUp_Timer()
  If Form Is Nothing Then Exit Sub
  If Form.WidgetRoot Is Nothing Then Exit Sub
  If Form.hWnd = 0 Then Exit Sub
  On Error Resume Next
    With Form.WidgetRoot
      If mMouseDown And .MouseKeyDown = 0 Then mMouseDown = False
      If (Not mMouseDown) And .MouseKeyDown Then
        If .GetWindowUnderCursor <> .hWnd Then RaiseEvent MouseClickOutside
      End If
    End With
    RaiseEvent CheckForAdditionalCloseConditions
  If Err Then Err.Clear
End Sub
 
Private Sub Class_Terminate()
  Unload
End Sub
