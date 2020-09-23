VERSION 5.00
Begin VB.Form frmOil 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   Picture         =   "frmOil.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer BallTimer 
      Interval        =   1
      Left            =   2520
      Top             =   4800
   End
   Begin VB.Image midh2 
      Height          =   270
      Left            =   4320
      Picture         =   "frmOil.frx":37C36
      Top             =   5400
      Width           =   1080
   End
   Begin VB.Image midh1 
      Height          =   270
      Left            =   4080
      Picture         =   "frmOil.frx":38BAA
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Image midv2 
      Height          =   1080
      Left            =   7080
      Picture         =   "frmOil.frx":39B1E
      Top             =   3360
      Width           =   270
   End
   Begin VB.Image midv1 
      Height          =   1080
      Left            =   3000
      Picture         =   "frmOil.frx":3AB22
      Top             =   3120
      Width           =   270
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Bttm 
      Height          =   270
      Left            =   4080
      Picture         =   "frmOil.frx":3BB26
      Top             =   8640
      Width           =   1080
   End
   Begin VB.Image Tp 
      Height          =   270
      Left            =   3960
      Picture         =   "frmOil.frx":3CA9A
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Lft 
      Height          =   1080
      Left            =   120
      Picture         =   "frmOil.frx":3DA0E
      Top             =   3240
      Width           =   270
   End
   Begin VB.Image Rght 
      Height          =   1080
      Left            =   11520
      Picture         =   "frmOil.frx":3EA12
      Top             =   3480
      Width           =   270
   End
   Begin VB.Image Ball 
      Height          =   375
      Left            =   4800
      Picture         =   "frmOil.frx":3FA16
      Top             =   6480
      Width           =   375
   End
End
Attribute VB_Name = "frmOil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ballBtm, ballRght As Integer
Dim Score As Integer

Private Sub BallTimer_Timer()
If strData.Dirc = DL Then
    Ball.Left = Ball.Left - levDat.Difficulty
    Ball.Top = Ball.Top + levDat.Difficulty
ElseIf strData.Dirc = DR Then
    Ball.Left = Ball.Left + levDat.Difficulty
    Ball.Top = Ball.Top + levDat.Difficulty
ElseIf strData.Dirc = UL Then
    Ball.Left = Ball.Left - levDat.Difficulty
    Ball.Top = Ball.Top - levDat.Difficulty
ElseIf strData.Dirc = UR Then
    Ball.Left = Ball.Left + levDat.Difficulty
    Ball.Top = Ball.Top - levDat.Difficulty
End If

If (Ball.Height + Ball.Top) >= Bttm.Top Then
    If (Ball.Left + Ball.Width) >= Bttm.Left And Ball.Left <= (Bttm.Left + Bttm.Width) And (Ball.Top + Ball.Height) < (Bttm.Top + 10) Then
        Score = Score + 1
        lblScore.Caption = "Score: " & Score
        If strData.Dirc = DR Then strData.Dirc = UR
        If strData.Dirc = DL Then strData.Dirc = UL
    End If
ElseIf Ball.Top <= (Tp.Top + Tp.Height) Then
    If (Ball.Left + Ball.Width) >= Tp.Left And Ball.Left <= (Tp.Left + Tp.Width) And Ball.Top > (Tp.Top - 10) Then
        Score = Score + 1
        lblScore.Caption = "Score: " & Score
        If strData.Dirc = UL Then strData.Dirc = DL
        If strData.Dirc = UR Then strData.Dirc = DR
    End If
ElseIf Ball.Left <= (Lft.Left + Lft.Left) Then
    If (Ball.Top + Ball.Height) >= Lft.Top And Ball.Top <= (Lft.Top + Lft.Height) And Ball.Left < (Lft.Left + Lft.Width - 10) Then
        Score = Score + 1
        lblScore.Caption = "Score: " & Score
        If strData.Dirc = UL Then strData.Dirc = UR
        If strData.Dirc = DL Then strData.Dirc = DR
    End If
ElseIf (Ball.Left + Ball.Width) >= Rght.Left Then
    If (Ball.Top + Ball.Height) >= Rght.Top And Ball.Top <= (Rght.Top + Rght.Height) And Ball.Left < (Rght.Left + 10) Then
        Score = Score + 1
        lblScore.Caption = "Score: " & Score
        If strData.Dirc = DR Then strData.Dirc = DL
        If strData.Dirc = UR Then strData.Dirc = UL
    End If
End If
For a = Ball.Top To (Ball.Top + Ball.Height)
For b = Ball.Left To (Ball.Left + Ball.Width)
For c = midv1.Top To (midv1.Top + midv1.Height)
For d = midv1.Left To (midv1.Left + midv1.Width)

Next
Next
Next
Next
If Ball.Top <= 0 Or Ball.Left <= 0 Or (Ball.Top + Ball.Height) >= frmOil.ScaleHeight Or (Ball.Left + Ball.Width) >= frmOil.ScaleWidth Then ShowCursor True: MsgBox "You ended with a score of: " & Score, vbOKOnly + vbInformation, "AVIN Message...": BallTimer.Enabled = False: frmStart.Show: Me.Hide: Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
BallTimer.Enabled = False
ShowCursor True: Result = MsgBox("Quit Level?", vbYesNo + vbCritical, "AVIN Message")
If Result = vbYes Then frmStart.Show: Me.Hide
If Result = vbNo Then BallTimer.Enabled = True: ShowCursor False
End Sub

Private Sub Form_Load()
ShowCursor False
Randomize Timer
strData.Dirc = Int(Rnd * 4) + 1
strData.Score = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
midv1.Top = y - (Lft.Height / 2)
midv2.Top = y - (Lft.Height / 2)
midh1.Left = x - (Tp.Width / 2)
midh2.Left = x - (Tp.Width / 2)
Tp.Left = x - (Tp.Width / 2)
Bttm.Left = x - (Tp.Width / 2)
Lft.Top = y - (Lft.Height / 2)
Rght.Top = y - (Lft.Height / 2)
End Sub

Private Sub Form_Terminate()
ShowCursor True
End Sub

