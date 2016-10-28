Attribute VB_Name = "Module1"
Type BallType
    xVel As Integer 'BALL VELOCITY
    yVel As Integer
    totalVel As Integer
            
End Type

Type PaddleType
    yVel As Integer
    AIControl As Boolean
    
End Type

Enum GameMode
    PvAI
    PvP
    AIvAI
    COOP

End Enum

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Global game As GameMode
Global friction As Single
Global scoreWin As Integer
Global ff As Integer 'FREEFILE
Global running As Boolean

Sub subCenterPos(obj As Control, frame As Form)

    obj.Left = frame.ScaleWidth / 2 - (obj.Width / 2)

End Sub

