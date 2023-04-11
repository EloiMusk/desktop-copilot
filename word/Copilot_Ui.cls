VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Copilot_Ui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Mycbar As CommandBar, toggleButton As CommandBarControl, updateStoreButton As CommandBarControl


Sub AddMenu()

Application.ScreenUpdating = False
RemoveMenu ' call remove routine to ensure only one menu in place

Set Mycbar = CommandBars.Add _
(Name:="Dektop Copilot", Position:=msoBarBottom, Temporary:=False)
' create new commandbar (menu bar)

Set updateStoreButton = Mycbar.Controls.Add(msoControlButton)
' create new commandbar control (button type) on custom menu
With updateStoreButton
.Caption = "Update Store" ' mouseover text
.Tag = "updateStore" ' used for identification
.OnAction = "updateStore" ' macro called with control
.FaceId = 6457 ' appearance, based on built-in faces
End With

Set toggleButton = Mycbar.Controls.Add(msoControlButton)
' create new commandbar control (button type) on custom menu
With toggleButton
.Caption = "Enable/Disable" ' mouseover text
.Tag = "toggle" ' used for identification
.OnAction = "toggle" ' macro called with control
.FaceId = 2997 ' appearance, based on built-in faces
End With

With Mycbar.Controls.Add(msoControlButton)
    .Caption = "Get Completion"
    .Tag = "getCompletion"
    .OnAction = "main"
    .FaceId = 7707
End With



Mycbar.Visible = True
Application.ScreenUpdating = True

Set Mycbar = Nothing 'release memory
Set toggleButton = Nothing

End Sub

Sub RemoveMenu()
Dim Mycbar As CommandBar

On Error Resume Next ' in case its already gone
Set Mycbar = CommandBars("Dektop Copilot")
Mycbar.Delete
Set Mycbar = Nothing 'release memory

End Sub

Private Sub Class_Initialize()
 AddMenu
End Sub
