Attribute VB_Name = "OmniPush"
Private Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Public Sub Omnipush_F8()
If GetKeyPress(vbKeyF1) Then
On Error Resume Next
Logger.Text2.Text = ""
Form2.Text3.Text = "0"
Form2.Text3.ForeColor = &H80FF80
F1:
End If
If GetKeyPress(vbKeyF1) Then
GoTo F1
End If
If GetKeyPress(vbKeyF2) Then
On Error Resume Next
SendKeys ("Susurrar") + (" ")
SendKeys "^" + ("{a}")
SendKeys "^" + ("{c}")
SendKeys "{BACKSPACE}"
F2:
End If
On Error Resume Next
If GetKeyPress(vbKeyF2) Then
GoTo F2
End If
Copy = Clipboard.GetText
If Copy = "" Then
Else
      If Form2.Text3.Text = "0" Then
      Logger.Text2.ForeColor = &H80FF80
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Logger.Text2.Text = Logger.sep.Text & " " & limp
      Logger.Text2.SelStart = Len(Logger.Text2.Text)
      Form2.Text3.Text = Form2.Text3.Text + 1
      Else
      Copy = Clipboard.GetText
      perla = Replace(Copy, "Susurrar ", "")
      limp = Replace(perla, " ", "")
      Logger.Text2.Text = Logger.Text2.Text & vbCrLf & Logger.sep.Text & " " & limp
      Logger.Text2.SelStart = Len(Logger.Text2.Text)
      Form2.Text3.Text = Form2.Text3.Text + 1
      End If
      Clipboard.Clear
If Form2.Text3.Text = "30" Then
Form2.Text3.ForeColor = &H80FFFF
Logger.Text2.ForeColor = &H80FFFF
End If
If Form2.Text3.Text = "60" Then
Form2.Text3.ForeColor = &H8080FF
Form2.adver.ForeColor = &H8080FF
Logger.Text2.ForeColor = &H8080FF
Form2.adver.Caption = " Excess"
End If
If Form2.Text3.Text = "100" Then
Form2.Text3.ForeColor = vbRed
Form2.adver.ForeColor = vbRed
Logger.Text2.ForeColor = vbRed
Form2.adver.Caption = " ¡Warning!"
End If
If Form2.Text3.Text = "300" Then
Form2.Text3.ForeColor = &HFF8080
Form2.adver.ForeColor = &HFF8080
Logger.Text2.ForeColor = &HFF8080
Form2.adver.Caption = " God level"
End If
If Form2.Text3.Text = "500" Then
Form2.Text3.ForeColor = &HC000C0
Form2.adver.ForeColor = &HC000C0
Logger.Text2.ForeColor = &HC000C0
Form2.adver.Caption = " Dione5G"
End If
If Form2.Text3.Text = "1000" Then
Form2.Text3.ForeColor = vbBlack
Form2.adver.ForeColor = vbBlack
Logger.Text2.ForeColor = vbBlack
Form2.adver.BackStyle = 1
Form2.adver.Caption = " DarkNet"
End If
End If
   End Sub
Public Sub OmnipushB()
If GetKeyPress(vbKeyF3) Then
On Error Resume Next
SendKeys (Logger.Text2.Text) + ("{enter}")
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
SendKeys (Form2.Combo1.Text) + ("{enter}")
F6:
End If
If GetKeyPress(vbKeyF6) Then
GoTo F6
End If
If GetKeyPress(vbKeyF8) Then
On Error Resume Next
Logger.Hide
Unload Logger
Form1.Show
Form2.Timer2.Enabled = False
Form2.Timer3.Enabled = False
Form2.Timer4.Enabled = False
Form2.Timer5.Enabled = False
Form2.Timer6.Enabled = False
Form2.omnipu.Enabled = False
Form1.Timer1.Enabled = True
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = True
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = True
Form2.Hide
Unload Form2
F8:
End If
If GetKeyPress(vbKeyF8) Then
GoTo F8
End If
End Sub
