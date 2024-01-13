Attribute VB_Name = "Push_Pull_M"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub Push_Pull()
If GetKeyPress(vbKeyF2) Then
On Error Resume Next
SendKeys (":push x") + ("{enter}")
F2:
End If
If GetKeyPress(vbKeyF2) Then
GoTo F2
End If
If GetKeyPress(vbKeyF3) Then
On Error Resume Next
SendKeys (":pull x") + ("{enter}")
F3:
End If
If GetKeyPress(vbKeyF3) Then
GoTo F3
End If
If GetKeyPress(vbKeyF4) Then
On Error Resume Next
SendKeys (":moonwalk") + ("{enter}")
F4:
End If
If GetKeyPress(vbKeyF4) Then
GoTo F4
End If
If GetKeyPress(vbKeyF5) Then
On Error Resume Next
SendKeys (":sit") + ("{enter}")
F5:
End If
If GetKeyPress(vbKeyF5) Then
GoTo F5
End If
If GetKeyPress(vbKeyF6) Then
On Error Resume Next
SendKeys (Form1.Combo1.Text) + ("{enter}")
F6:
End If
If GetKeyPress(vbKeyF6) Then
GoTo F6
End If
If GetKeyPress(vbKeyF8) Then
On Error Resume Next
Form2.Show
Logger.Show
Logger.Left = Form2.Left
Logger.Top = Form2.Top + Form2.Height
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = False
Form1.Hide
Form1.Timer1.Enabled = False
F8:
End If
If GetKeyPress(vbKeyF8) Then
GoTo F8
End If
End Sub
