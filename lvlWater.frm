VERSION 5.00
Begin VB.Form frmSpace 
   BorderStyle     =   0  'None
   Caption         =   "Pong!"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "lvlSpace.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer BallTimer 
      Interval        =   1
      Left            =   2640
      Top             =   4920
   End
   Begin VB.Image Ball 
      Height          =   375
      Left            =   5520
      Picture         =   "lvlSpace.frx":341BE
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image Rght 
      Height          =   1080
      Left            =   11640
      Picture         =   "lvlSpace.frx":345D5
      Top             =   3600
      Width           =   270
   End
   Begin VB.Image Lft 
      Height          =   1080
      Left            =   120
      Picture         =   "lvlSpace.frx":355D9
      Top             =   3360
      Width           =   270
   End
   Begin VB.Image Tp 
      Height          =   270
      Left            =   4080
      Picture         =   "lvlSpace.frx":365DD
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Bttm 
      Height          =   270
      Left            =   4200
      Picture         =   "lvlSpace.frx":37551
      Top             =   8640
      Width           =   1080
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "frmSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Score As Integer

Private Sub BallTimer_Timer()
If strData.Dirc = DL Then
    Ball.Left = Ball.Left - 7
    Ball.Top = Ball.Top + 7
ElseIf strData.Dirc = DR Then
    Ball.Left = Ball.Left + 7
    Ball.Top = Ball.Top + 7
ElseIf strData.Dirc = UL Then
    Ball.Left = Ball.Left - 7
    Ball.Top = Ball.Top - 7
ElseIf strData.Dirc = UR Then
    Ball.Left = Ball.Left + 7
    Ball.Top = Ball.Top - 7
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
If Ball.Top <= 0 Or Ball.Left <= 0 Or (Ball.Top + Ball.Height) >= frmSpace.ScaleHeight Or (Ball.Left + Ball.Width) >= frmSpace.ScaleWidth Then ShowCursor True: MsgBox "You ended with a score of: " & Score, vbOKOnly + vbInformation, "AVIN Message...": BallTimer.Enabled = False: frmStart.Show: Me.Hide: Unload Me
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
Tp.Left = x - (Tp.Width / 2)
Bttm.Left = x - (Tp.Width / 2)
Lft.Top = y - (Lft.Height / 2)
Rght.Top = y - (Lft.Height / 2)
End Sub

Private Sub Form_Terminate()
ShowCursor True
End Sub
