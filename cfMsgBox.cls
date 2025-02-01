VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private MyText As cwTextBox

Private WithEvents Button1 As cwButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2 As cwButton
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3 As cwButton
Attribute Button3.VB_VarHelpID = -1

Private BOXWidth As Double
Private BOXHeight As Double

Private Style  As VbMsgBoxStyle

Private IconName As Long

Public ReturnValue As VbMsgBoxResult

Private Zoom   As Double


Private Sub Button1_Click()

    Select Case Style

    Case vbInformation
        ReturnValue = vbOK
        Form.Unload
    Case vbOKOnly
        ReturnValue = vbOK
        Form.Unload
    Case vbExclamation
        ReturnValue = vbOK
        Form.Unload
    Case Else
        MsgBox "not implemented yet"
    End Select

End Sub

Private Sub Class_Initialize()
 
    Zoom = 1 'fMain.CurrZoom
    If Zoom = 0 Then Zoom = 1

    BOXWidth = 500 * Zoom
    BOXHeight = 300 * Zoom

    Set Form = Cairo.WidgetForms.Create(vbFixedSingle, "Caption", True, BOXWidth, BOXHeight)
    Form.Hide
    Form.WidgetRoot.Zoom = Zoom    '.............<<<<<<<<< Important
    Form.IconImageKey = "Icon4"
    Form.WindowState = vbNormal

    Form.WidgetRoot.BackColor = &HBBBBBB    ' &H8000000F
    Form.CenterOn New_c.Displays(1)


End Sub


Public Function DisplayMessage(Message As String, Optional Title As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
    Dim TEXTHeight As Double
    Dim ButtonWidth As Double
    Dim ButtonHeight As Double

    Dim BOXLeft As Double
    Dim BOXTop As Double


    Style = Buttons


    fMain.Form.Enabled = False


    ButtonHeight = Zoom * 30    ' 22
    ButtonWidth = Zoom * 120    '80


    Form.Caption = Title


    If Form.Widgets.Exists("TB") Then
        Form.Widgets.Remove "TB"
    End If

    '    Set MyText = W.Widgets.Add(New cwTextBox, "TB", 20, 10, BOXWidth - 30, BOXHeight - 50)
    '    Set MyText = Form.Widgets.Add(New cwTextBox, "TB", 50, 15, BOXWidth - 60, BOXHeight - 50)

    Set MyText = Form.Widgets.Add(New cwTextBox, "TB", 30 * Zoom, 15 * Zoom, BOXWidth - 60 * Zoom, BOXHeight - 50 * Zoom)


    With MyText

        .Widget.BackColor = &HBBBBBB    'Form.BackColor
        .Widget.FontSize = 10.5
        .RowHeight = .Widget.FontSize * 1.55

        .Text = Message
        .Border = False
        .MultiLine = True

        .EnsureVisible
        .Locked = True
 
        TEXTHeight = .RowHeight * (.RowCount + 1)

        If TEXTHeight > 430 / Zoom Then TEXTHeight = 430 / Zoom

        'Form.Height = TEXTHeight + 40 + ButtonHeight
        BOXHeight = TEXTHeight + ButtonHeight * 5

        .Widget.Move 30 / Zoom, 15 / Zoom, .Widget.Width / Zoom, TEXTHeight / Zoom

    End With

    Select Case Buttons

    Case vbInformation

        ' Set Button1 = W.Widgets.Add(New cwButton, "Button1", BOXWidth * 0.5 - ButtonWidth * 0.5, BOXHeight - ButtonHeight - 6, ButtonWidth, ButtonHeight)
        '        Set Button1 = Form.Widgets.Add(New cwButton, "Button1", BOXWidth * 0.5 - ButtonWidth * 0.5, BOXHeight - ButtonHeight * 3, ButtonWidth, ButtonHeight)
        Set Button1 = Form.Widgets.Add(New cwButton, "Button1", (BOXWidth * 0.5) / Zoom - ButtonWidth * 0.5, (TEXTHeight + ButtonHeight * 2) / Zoom, ButtonWidth, ButtonHeight)

        Set Button2 = Nothing
        Set Button3 = Nothing
        Button1.Caption = "OK"
        Button1.Widget.MousePointer = IDC_HAND

        IconName = 33

    Case vbOKOnly

        'Set Button1 = W.Widgets.Add(New cwButton, "Button1", BOXWidth * 0.5 - ButtonWidth * 0.5, BOXHeight - ButtonHeight - 5, ButtonWidth, ButtonHeight)
        Set Button1 = Form.Widgets.Add(New cwButton, "Button1", (BOXWidth * 0.5) / Zoom - ButtonWidth * 0.5, (TEXTHeight + ButtonHeight * 2) / Zoom, ButtonWidth, ButtonHeight)
        Set Button2 = Nothing
        Set Button3 = Nothing
        Button1.Caption = "OK"
        Button1.Widget.MousePointer = IDC_HAND

    Case vbExclamation
        Set Button1 = Form.Widgets.Add(New cwButton, "Button1", (BOXWidth * 0.5) / Zoom - ButtonWidth * 0.5, (TEXTHeight + ButtonHeight * 2) / Zoom, ButtonWidth, ButtonHeight)
        Set Button2 = Nothing
        Set Button3 = Nothing
        Button1.Caption = "OK"
        MyText.Widget.FontBold = "TRUE"
        MyText.Alignment = vbCenter
        Button1.Widget.MousePointer = IDC_HAND

    Case Else
'        DisplayMessage = "Not Implemented yet"
    End Select



    BOXLeft = (fMain.Form.Width) * 0.5 * 1 - BOXWidth * 0.5
    BOXTop = (fMain.Form.Height) * 0.5 * 1 - BOXHeight * 0.5

    Form.Move BOXLeft, BOXTop, BOXWidth, BOXHeight

    Form.CenterOn New_c.Displays(1)
    Form.Show 1
 
End Function




Private Sub Form_KeyPress(KeyUnicode As Integer)

    If KeyUnicode = 13 Then
        ReturnValue = vbOK
        Form.Unload
    End If

End Sub

Private Sub Form_Paint()
'  Cairo.ImageList.AddIconFromResourceFile "Icon" & CStr(IconName), "shell32.dll", IconName, 32, 32

End Sub

Private Sub Form_Unload(Cancel As Integer)
    fMain.Form.Enabled = True
End Sub

Private Sub Class_Terminate()
    Set MyText = Nothing
    Set Button1 = Nothing
    Set Button2 = Nothing
    Set Button3 = Nothing
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)

'
    CC.SetSourceColor &HBBBBBB
    CC.Rectangle 0, 0, dx_Aligned, dy_Aligned
    CC.Fill
    '
    '    If Not (Button1 Is Nothing) Then Button1.Widget.Refresh
    '  '  If Not (MyText Is Nothing) Then MyText.Widget.Refresh

End Sub
