VERSION 5.00
Begin VB.Form BreakoutFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   240
      Top             =   8640
   End
   Begin VB.Label lblpause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press Space to pause or continue"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label lblscore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label lblscore2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   7680
      Width           =   855
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   29
      Left            =   11880
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   28
      Left            =   10560
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   9240
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   26
      Left            =   7920
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   25
      Left            =   6600
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   5280
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   23
      Left            =   3960
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   22
      Left            =   2640
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   1320
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   0
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   19
      Left            =   11880
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   18
      Left            =   10560
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   9240
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   16
      Left            =   7920
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   15
      Left            =   6600
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   5280
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   3960
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   12
      Left            =   2640
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   11
      Left            =   1320
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   0
      Top             =   840
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   9
      Left            =   11880
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   10560
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   9240
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   7920
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   5
      Left            =   6600
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   5280
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   3960
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   2640
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   1320
      Top             =   240
      Width           =   1200
   End
   Begin VB.Shape shpbrick 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lbllives2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lives"
      Height          =   375
      Left            =   10920
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lbllives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   11040
      TabIndex        =   0
      Top             =   7920
      Width           =   975
   End
   Begin VB.Shape shpbat 
      BorderColor     =   &H00000008&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Top             =   8640
      Width           =   4815
   End
   Begin VB.Shape shpball 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "BreakoutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bricks(2, 9) As Integer
Dim bounceleft As Integer
Dim bouncetop As Integer
Dim ball As Integer
Dim middle As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyLeft
        shpbat.Left = shpbat.Left - 150
    Case vbKeyRight
        shpbat.Left = shpbat.Left + 150
    Case vbKeySpace
        If Timer1.Enabled = True Then
            Timer1.Enabled = False
        Else
            If Timer1.Enabled = False Then
                Timer1.Enabled = True
            End If
        End If
    Case 65
        Call reset
End Select
End Sub

Private Sub Form_Load()
shpball.Left = BreakoutFrm.ScaleWidth / 2
shpball.Top = BreakoutFrm.ScaleHeight / 2
bouncetop = 40
bounceleft = 40
lives = 5
score = 0
lbllives = lives
lblscore = score
For Counter = 0 To 9
    bricks(0, Counter) = 0
    bricks(1, Counter) = 0
    bricks(2, Counter) = 0
Next
End Sub

Private Sub brick()
For Counter = 0 To 29
    If bricks((Counter \ 10), (Counter Mod 10)) = 0 And shpball.Top > (shpbrick(Counter).Top + shpbrick(Counter).Height) - 25 And shpball.Top < (shpbrick(Counter).Top + shpbrick(Counter).Height) + 25 And middle >= shpbrick(Counter).Left And middle <= shpbrick(Counter).Left + shpbrick(Counter).Width Then
        bouncetop = bouncetop * -1
        bounceleft = bounceleft * -1
        bricks((Counter \ 10), (Counter Mod 10)) = 1
        shpbrick(Counter).Visible = False
        lblscore = lblscore + 1
    End If
    If bricks((Counter \ 10), (Counter Mod 10)) = 0 And (shpball.Top + shpball.Height) > shpbrick(Counter).Top - 25 And (shpball.Top + shpball.Height) < shpbrick(Counter).Top + 25 And middle >= shpbrick(Counter).Left And middle <= shpbrick(Counter).Left + shpbrick(Counter).Width Then
        bouncetop = bouncetop * -1
        bounceleft = bounceleft * -1
        bricks((Counter \ 10), (Counter Mod 10)) = 1
        shpbrick(Counter).Visible = False
        lblscore = lblscore + 1
    End If
    If bricks((Counter \ 10), (Counter Mod 10)) = 0 And shpball.Left > (shpbrick(Counter).Left + shpbrick(Counter).Width) - 25 And shpball.Left < (shpbrick(Counter).Left + shpbrick(Counter).Width) + 25 And (shpball.Top + (shpball.Height / 2)) > shpbrick(Counter).Top And (shpball.Top + (shpball.Height / 2)) < (shpbrick(Counter).Top + shpbrick(Counter).Height) Then
        bouncetop = bouncetop * -1
        bounceleft = bounceleft * -1
        bricks((Counter \ 10), (Counter Mod 10)) = 1
        shpbrick(Counter).Visible = False
        lblscore = lblscore + 1
    End If
    If bricks((Counter \ 10), (Counter Mod 10)) = 0 And (shpball.Left + shpball.Width) > shpbrick(Counter).Left - 25 And (shpball.Left + shpball.Width) < shpbrick(Counter).Left + 25 And (shpball.Top + (shpball.Height / 2)) > shpbrick(Counter).Top And (shpball.Top + (shpball.Height / 2)) < (shpbrick(Counter).Top + shpbrick(Counter).Height) Then
        bouncetop = bouncetop * -1
        bounceleft = bounceleft * -1
        bricks((Counter \ 10), (Counter Mod 10)) = 1
        shpbrick(Counter).Visible = False
        lblscore = lblscore + 1
    End If
Next
End Sub

Private Sub Timer1_Timer()
shpball.Top = shpball.Top + bouncetop
shpball.Left = shpball.Left + bounceleft
score = lblscore
If shpball.Left >= (BreakoutFrm.ScaleWidth - shpball.Width) Then
    bounceleft = bounceleft * -1
End If
If shpball.Top >= (BreakoutFrm.ScaleHeight - shpball.Height) Then
    lbllives = lbllives - 1
    shpball.Left = BreakoutFrm.ScaleWidth / 2
    shpball.Top = BreakoutFrm.ScaleHeight / 2
    Timer1.Enabled = False
    If lbllives = 0 Then
        MsgBox ("Game Over")
        MsgBox ("You scored " & score & " Points")
        Call reset
    End If
End If
If shpball.Left <= 0 Then
    bounceleft = bounceleft * -1
End If
If shpball.Top <= 0 Then
    bouncetop = bouncetop * -1
End If
middle = shpball.Left + (shpball.Width / 2)
ball = shpball.Top + shpball.Height
If ball > shpbat.Top - 50 And ball < shpbat.Top + 50 Then
    If middle >= shpbat.Left + 250 And middle <= (shpbat.Left + (shpbat.Width / 2)) Then
        bouncetop = bouncetop * -1
        bounceleft = bounceleft * -1
    Else
        If middle >= (shpbat.Left + (shpbat.Width / 2)) And middle < ((shpbat.Left + shpbat.Width) - 250) Then
            bouncetop = bouncetop * -1
        Else
            If middle >= ((shpbat.Left + shpbat.Width) - 250) And middle <= (shpbat.Left + shpbat.Width) Then
                bouncetop = bouncetop * -1
                bounceleft = bounceleft + 70
            Else
                If middle >= shpbat.Left And middle <= shpbat.Left + 250 Then
                    bouncetop = bouncetop * -1
                    bounceleft = (bounceleft + 70) * -1
                End If
            End If
        End If
    End If
End If
Call brick
If score = 30 Then
    Call reset
End If
End Sub
