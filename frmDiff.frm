VERSION 5.00
Begin VB.Form frmDiff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVIN Four Paddle Pong"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   1935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "<< Back"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Your Difficulty"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.CommandButton btnHard4 
         Caption         =   "Impossible"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton btnHard3 
         Caption         =   "Hard"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton btnHard2 
         Caption         =   "Easy"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton btnHard1 
         Caption         =   "Very Easy"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
frmLevel.Show
Me.Hide
End Sub

Private Sub btnHard1_Click()
levDat.Difficulty = 7
setLevel levDat.Level
End Sub

Private Sub btnHard2_Click()
levDat.Difficulty = 10
setLevel levDat.Level
End Sub

Private Sub btnHard3_Click()
levDat.Difficulty = 13
setLevel levDat.Level
End Sub

Private Sub btnHard4_Click()
levDat.Difficulty = 15
setLevel levDat.Level
End Sub

Private Sub btnQuit_Click()
Result = MsgBox("Are you sure you want to quit?", vbYesNo + vbCritical, "AVIN Message...")
If Result = vbYes Then End
End Sub
