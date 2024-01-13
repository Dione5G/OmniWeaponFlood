Attribute VB_Name = "Turbo_Push_F7"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub Turbo_Push()
If GetKeyPress(vbKeyF7) Then
On Error Resume Next
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
SendKeys (":push x") + ("{enter}")
F7:
End If
If GetKeyPress(vbKeyF7) Then
GoTo F7
End If
End Sub
