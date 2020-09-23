Attribute VB_Name = "Declares"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Type TypA
    Dirc As Integer
    Score As Integer
End Type
Type TypB
    Level As Integer
    Difficulty As Integer
End Type
Public strData As TypA
Public Const DL = 1
Public Const DR = 2
Public Const UL = 3
Public Const UR = 4
Public Result As VbMsgBoxResult
Public levDat As TypB

Public Sub setLevel(theLevel As Integer)
If theLevel = 1 Then frmSpace.Show
If theLevel = 2 Then frmOil.Show
frmDiff.Hide
End Sub
