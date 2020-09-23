VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVIN Four Paddle Pong"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnHighScore 
      Caption         =   "&High Scores"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnPick 
      Caption         =   "&Pick a Level"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnHighScore_Click()
MsgBox "This feature is unavailable at this time.", vbOKOnly + vbExclamation, "AVIN Message..."
End Sub

Private Sub btnPick_Click()
frmLevel.Show
Me.Hide
End Sub

Private Sub btnQuit_Click()
Result = MsgBox("Are you sure you want to quit?", vbYesNo + vbCritical, "AVIN Message...")
If Result = vbYes Then End
End Sub
