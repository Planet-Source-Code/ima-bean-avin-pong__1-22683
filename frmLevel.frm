VERSION 5.00
Begin VB.Form frmLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVIN Four Paddle Pong"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "<< Back"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame framePick 
      Caption         =   "Pick a Level"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton btnOil 
         Caption         =   "Oil Tankers"
         Height          =   1455
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton btnSpace 
         Caption         =   "Space"
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()
frmStart.Show
Me.Hide
End Sub

Private Sub btnOil_Click()
If MsgBox("This form is very big and could cause your computer to freeze." & vbNewLine & "It is also not finished. Do you want to continue?", vbYesNo + vbQuestion, "Continue?") = vbYes Then
    levDat.Level = 2
    frmDiff.Show
    Me.Hide
End If
End Sub

Private Sub btnQuit_Click()
Result = MsgBox("Are you sure you want to quit?", vbYesNo + vbCritical, "AVIN Message...")
If Result = vbYes Then End
End Sub

Private Sub btnSpace_Click()
levDat.Level = 1
frmDiff.Show
Me.Hide
End Sub
