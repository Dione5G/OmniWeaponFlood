Attribute VB_Name = "AutloClic"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSELEFTDOWN = &H2
Private Const MOUSELEFTUP = &H4
Public Sub autoclick()
If GetKeyPress(vbKeyF9) Then
On Error Resume Next
Form1.fx9.ForeColor = vbGreen
Form1.fx10.ForeColor = &HFFFFFF
Form1.Timer2.Enabled = True
Form1.Timer4.Enabled = False
Form1.Shape3.Visible = True
Form1.Shape4.Visible = False
F9:
End If
If GetKeyPress(vbKeyF9) Then
GoTo F9
End If
If GetKeyPress(vbKeyF10) Then
On Error Resume Next
Form1.fx9.ForeColor = &HFFFFFF
Form1.fx10.ForeColor = vbGreen
Form1.Timer2.Enabled = False
Form1.Timer4.Enabled = True
Form1.Shape3.Visible = True
Form1.Shape4.Visible = False
F10:
End If
If GetKeyPress(vbKeyF10) Then
GoTo F10
End If
If GetKeyPress(vbKeyF11) Then
On Error Resume Next
Form1.Timer2.Enabled = False
Form1.Timer4.Enabled = False
Form1.Shape3.Visible = False
Form1.Shape4.Visible = True
Form1.fx9.ForeColor = &HFFFFFF
Form1.fx10.ForeColor = &HFFFFFF
F11:
End If
If GetKeyPress(vbKeyF11) Then
GoTo F11
End If
End Sub
Public Sub oneclic()
mouse_event MOUSELEFTDOWN, 0, 0, 0, 0
mouse_event MOUSELEFTUP, 0, 0, 0, 0
End Sub
Public Sub dobleclic()
mouse_event MOUSELEFTDOWN, 0, 0, 0, 0
mouse_event MOUSELEFTUP, 0, 0, 0, 0
mouse_event MOUSELEFTDOWN, 0, 0, 0, 0
mouse_event MOUSELEFTUP, 0, 0, 0, 0
End Sub
Public Sub autoclickOmni()
If GetKeyPress(vbKeyF9) Then
On Error Resume Next
Form2.fx9.ForeColor = vbGreen
Form2.fx10.ForeColor = &HFFFFFF
Form2.Timer2.Enabled = True
Form2.Timer4.Enabled = False
Form2.Shape3.Visible = True
Form2.Shape4.Visible = False
F9:
End If
If GetKeyPress(vbKeyF9) Then
GoTo F9
End If
If GetKeyPress(vbKeyF10) Then
On Error Resume Next
Form2.fx10.ForeColor = vbGreen
Form2.fx9.ForeColor = &HFFFFFF
Form2.Timer2.Enabled = False
Form2.Timer4.Enabled = True
Form2.Shape3.Visible = True
Form2.Shape4.Visible = False
F10:
End If
If GetKeyPress(vbKeyF10) Then
GoTo F10
End If
If GetKeyPress(vbKeyF11) Then
On Error Resume Next
Form2.Timer2.Enabled = False
Form2.Timer4.Enabled = False
Form2.Shape3.Visible = False
Form2Shape4.Visible = True
Form2.fx9.ForeColor = &HFFFFFF
Form2.fx10.ForeColor = &HFFFFFF
F11:
End If
If GetKeyPress(vbKeyF11) Then
GoTo F11
End If
End Sub
