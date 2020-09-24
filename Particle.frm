VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Particles - 0fps - 0 particles"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1800
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FPS As Integer
Dim Xx As Integer
Dim Yx As Integer
Dim Particle(0 To 1000) As Part
Dim NumPart As Integer
Dim Wind As Integer
Dim Gravity As Integer
Dim Refreshh As Boolean
Dim Strength As Integer
Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Particles                  '
'  Customize the Variables here '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
NumPart = 150 'How many particles?
Refreshh = False 'Refresh?
Gravity = 1 'How much gravity?
Wind = 0  ' How much wind?
Strength = 10 'How powerful emission can it get to
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tips;;;;;;;'
';;;;;;;;;;;;;'
'Set Gravity to 0 for starfield
'Set Wind to 10 or -10 for Comet
'Set Stength to 300 for wierd star
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Xx = x
Yx = y
End Sub

Private Sub Timer1_Timer()
FPS = FPS + 1
For i = 0 To NumPart
If Particle(i).active = True Then
Else
Particle(i).x = Xx
Particle(i).y = Yx
Particle(i).yf = Int(Rnd * Strength) - Int(Rnd * Strength)
Particle(i).xf = Int(Rnd * Strength) - Int(Rnd * Strength)
Particle(i).active = True
Particle(i).color = Rnd * vbWhite
End If
Next
If Refreshh = True Then Me.Refresh
For i = 0 To NumPart
If Particle(i).active = True Then
Particle(i).ox = Particle(i).x
Particle(i).oy = Particle(i).y
Particle(i).x = Particle(i).x + Particle(i).xf
Particle(i).y = Particle(i).y + Particle(i).yf
Particle(i).yf = Particle(i).yf + Gravity
Particle(i).xf = Particle(i).xf + Wind
If Particle(i).y > Me.ScaleHeight Or Particle(i).x > Me.ScaleWidth Or Particle(i).y < 0 Or Particle(i).x < 0 Then
Particle(i).active = False
End If
Me.Line (Particle(i).ox, Particle(i).oy)-(Particle(i).x, Particle(i).y), Particle(i).color
End If
Next
End Sub

Private Sub Timer2_Timer()
Me.Caption = "Particles - FPS:" & FPS & " - Particles:" & NumPart
FPS = 0
End Sub
