VERSION 5.00
Begin VB.Form formGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".pong"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWinName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   100
      Left            =   600
      Top             =   4920
   End
   Begin VB.Timer tmrPause 
      Interval        =   1000
      Left            =   120
      Top             =   840
   End
   Begin VB.Timer tmrMain 
      Interval        =   10
      Left            =   120
      Top             =   240
   End
   Begin VB.Label lblEnter 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game start!"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   10935
   End
   Begin VB.Shape circBall 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   180
   End
   Begin VB.Line lineMid 
      BorderStyle     =   2  'Dash
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   6480
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   750
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   6120
      Width           =   10935
   End
   Begin VB.Label lblPaddle 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1000
      Index           =   1
      Left            =   9720
      TabIndex        =   1
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblPaddle 
      BackColor       =   &H00000000&
      Height          =   1000
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   255
   End
End
Attribute VB_Name = "formGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Yes, I used to name variables and functions in extended Hungarian notation

Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const PI = 3.14159265359
Const MAX_ANGLE = 5 * PI / 12 '75 DEGREES IN RADIANS
Const MAX_PADV = 200 'MAX PADDLE SPEED
Const INIT_PADV = 25 'ALSO RATE OF ACCELERATION PER INTERVAL
Const INIT_BALLV = 150

Dim ball As BallType
Dim pad(1) As PaddleType '0 TO 1 FOR EACH PLAYER

Dim dblAngle As Double
Dim intScore(1) As Integer

Dim intLastHit As Integer

Dim boolSwitch As Boolean

Dim intWinner As Integer

Dim i As Integer

Private Sub subInsertSort(arr, tagged As Boolean, Optional boundVal)
    
    Dim temp
    Dim k As Integer
    
    If tagged Then
        Dim bondTemp
    
    End If
    
    For i = 1 To UBound(arr)
        If tagged Then
            bondTemp = boundVal(i)
            
        End If
        
        temp = arr(i)
                
        k = i
        
        Do While temp > arr(k - 1)
            arr(k) = arr(k - 1)
            
            If tagged Then
                boundVal(k) = boundVal(k - 1)
            
            End If
            
            k = k - 1
            
            If k = 0 Then
                Exit Do
            
            End If
        
        Loop
        
        arr(k) = temp
        
        If tagged Then
            boundVal(k) = bondTemp
        
        End If
    
    Next i

End Sub

Private Function fnCheckCollision(objA As Control, objB As Control) As Boolean

    'CHECKS FOR COLLISION
    fnCheckCollision = Not ((objA.Top + objA.Height < objB.Top) Or (objB.Top + objB.Height < objA.Top) Or _
    (objA.Left + objA.Width < objB.Left) Or (objB.Left + objB.Width < objA.Left))
    
    If fnCheckCollision = True Then
        sndPlaySound App.Path & "\thunk.wav", &H80 Or &H1
        
    End If
    
End Function

Private Function fnCenter(obj As Control) As Integer
    
    fnCenter = obj.Top + (obj.Height / 2)

End Function

Private Function fnUpdtLeg(Leg2 As Integer, Hyp As Integer) As Integer 'A^2 + B^2 = C^2 PYTHAGOREAN THEOREM

    fnUpdtLeg = Int(Sqr((Hyp ^ 2) - (Leg2 ^ 2)))

End Function

Private Sub subUpdtAngle()

    If Not ball.xVel = 0 Then
        dblAngle = Atn(ball.yVel / ball.xVel)
        
    End If

End Sub

Private Sub subAI()

    For i = 0 To 1
        If pad(i).AIControl = True Then
            
            If intLastHit <> i And intLastHit <> -1 Then
                Call subAimMovePad(Int(i), circBall)

            Else
                If fnCenter(lblPaddle(i)) < Me.ScaleHeight / 2 - (MAX_PADV * 5) Then
                    Call subAccelPad(Int(i), -INIT_PADV)
                    
                ElseIf fnCenter(lblPaddle(i)) > Me.ScaleHeight / 2 + (MAX_PADV * 5) Then
                    Call subAccelPad(Int(i), INIT_PADV)
                    
                End If
                
            End If
            
        End If
        
    Next i

End Sub

Private Sub subAccelPad(PadI As Integer, Accel As Integer)

    If (pad(PadI).yVel < 0 And Accel > 0) Or (pad(PadI).yVel > 0 And Accel < 0) Then
        pad(PadI).yVel = pad(PadI).yVel + (2 * Accel)
                    
    End If
        
    If pad(PadI).yVel < MAX_PADV And pad(PadI).yVel > -MAX_PADV Then
        pad(PadI).yVel = pad(PadI).yVel + Accel
        
    End If

End Sub

Private Sub subAimMovePad(PadI As Integer, Target As Control)

    If fnCenter(lblPaddle(PadI)) > fnCenter(Target) And pad(PadI).yVel < MAX_PADV And (Not fnPadHitWall(PadI)) Then
        Call subAccelPad(PadI, INIT_PADV)
                    
    ElseIf fnCenter(lblPaddle(PadI)) < fnCenter(Target) And _
    pad(PadI).yVel > -MAX_PADV And _
    Not fnPadHitWall(PadI) Then
        Call subAccelPad(PadI, -INIT_PADV)
                    
    End If

End Sub

Private Sub subLimAngle()

    If Abs(Atn(ball.yVel / ball.xVel)) >= MAX_ANGLE Then
        If ball.xVel > 0 Then
            ball.xVel = Cos(MAX_ANGLE) * ball.totalVel
            
        Else
            ball.xVel = Cos(MAX_ANGLE) * ball.totalVel
            
        End If
            
        If ball.yVel > 0 Then
            ball.yVel = fnUpdtLeg(ball.xVel, ball.totalVel)
            
        Else
            ball.yVel = -fnUpdtLeg(ball.xVel, ball.totalVel)
            
        End If
        
        ball.totalVel = Sqr((ball.xVel ^ 2) + (ball.yVel ^ 2))
                    
    End If

End Sub

Private Sub subHitPad()
    
    For i = 0 To 1
        If fnCheckCollision(circBall, lblPaddle(i)) Then
            lblPaddle(i).BackColor = vbYellow
            
            If game = COOP And instlasthit <> i Then
                Call subCoopScore
            
            End If
            
            intLastHit = i
            
            
            'If (circBall.Left > lblPaddle(0).Left + lblPaddle(0).Width - Abs(ball.xVel)) And _
            '(circBall.Left + circBall.Width < lblPaddle(1).Left + Abs(ball.xVel)) Then
                
            If circBall.Left + circBall.Width < lblPaddle(0).Left + lblPaddle(0).Width Or _
            circBall.Left > lblPaddle(1).Left Then
                ball.yVel = -ball.yVel '+ (Round(pad(i).yVel * 0.5))

                ball.totalVel = Sqr((ball.xVel ^ 2) + (ball.yVel ^ 2))
            
            Else
                With ball
                    .yVel = .yVel + (Round(pad(i).yVel * friction))
                    
                    dblAngle = Atn(ball.yVel / -ball.xVel)
                    
                    If (.xVel < 0 And i = 0) Or (.xVel > 0 And i = 1) Then
                        .xVel = -.xVel
                        
                    End If
                    
                    
                    If Sqr((.xVel ^ 2) + (.yVel ^ 2)) < .totalVel Then 'MINIMUM SPEED
                        If .xVel >= 0 Then
                            .xVel = Round(Cos(dblAngle) * .totalVel, 0)

                        Else
                            .xVel = -Round(Cos(dblAngle) * .totalVel, 0)

                        End If
                        
                        If .yVel >= 0 Then
                            .yVel = fnUpdtLeg(.xVel, .totalVel)
                            
                        Else
                            .yVel = -fnUpdtLeg(.xVel, .totalVel)
                            
                        End If
                                            
                    End If
                    
                    .totalVel = Sqr((.xVel ^ 2) + (.yVel ^ 2))
                    
                    Call subLimAngle
                    
                End With
                
                
            'Else 'BOUNCE OFF TOP/BOTTOM
                'ball.yVel = -ball.yVel '+ (Round(pad(i).yVel * 0.5))

                'ball.totalVel = Sqr((ball.xVel ^ 2) + (ball.yVel ^ 2))
            
            End If
            
        End If
            
    Next i

End Sub

Private Function subHitWall()

    If (circBall.Top <= 0 And ball.yVel > 0) Then
        ball.yVel = -ball.yVel
        
        sndPlaySound App.Path & "\weakThunk.wav", &H80 Or &H1
        
    ElseIf (circBall.Top + circBall.Height >= Me.ScaleHeight And ball.yVel < 0) Then
        ball.yVel = -ball.yVel
        
        sndPlaySound App.Path & "\weakThunk.wav", &H80 Or &H1
        
    End If
        
    Call subUpdtAngle
   
End Function

Private Sub subLose()
    
    Call subStop
    
    lblPrompt.Caption = "Mission failed."
    lblPrompt.Visible = True
    lblEnter.Caption = "Continue"
    lblEnter.Visible = True

End Sub

Private Sub subCheckBounds()

    If circBall.Left <= 0 Then
        If game = COOP Then
            Call subLose
                
        Else
            Call subScored(1)
    
        End If
        
    ElseIf circBall.Left + circBall.Width >= Me.ScaleWidth Then
        If game = COOP Then
            Call subLose
                
        Else
            Call subScored(0)
    
        End If
        
    End If

End Sub

Private Sub subCoopScore()

    lblScore(0).Caption = lblScore(0).Caption + 1
    
    If lblScore(0).Caption = scoreWin Then
        Call subWin
    End If


End Sub

Private Sub subScored(PadI As Integer)
    
    intScore(PadI) = intScore(PadI) + 1
    
    
    lblScore(PadI).Caption = intScore(PadI)

    
    If intScore(PadI) = scoreWin Then
        intWinner = PadI
        Call subWin
        
    Else
        Select Case PadI
            Case 0
                circBall.Left = lblPaddle(1).Left - circBall.Width - 1
                intLastHit = 1
            
            Case 1
                circBall.Left = lblPaddle(0).Left + lblPaddle(0).Width + 1
                intLastHit = 0
                
        End Select
            
        ball.xVel = 0
        ball.yVel = 0
        ball.totalVel = 0
        
        lblPrompt.Caption = "Player " & (PadI + 1) & " scored!"
        lblPrompt.Visible = True
        
        tmrPause.Enabled = True
        
    End If
    
End Sub

Private Sub subKeys(CheckKey As Long, PadI As Integer, Accel As Integer) 'KEY FOR VBKEY VALUE AND PADI FOR INDEX, ACCEL FOR VALUE OF ACCELERATION
    
    If GetAsyncKeyState(CheckKey) And ((Accel > 0 And lblPaddle(PadI).Top > 0) Or (Accel < 0 And lblPaddle(PadI).Top + lblPaddle(PadI).Height < Me.ScaleHeight)) And _
    Not fnPadHitWall(PadI) Then 'LIMIT MAXIMUM VELOCITY
        Call subAccelPad(PadI, Accel)

    Else
        If (Accel > 0 And pad(PadI).yVel > 0) Or (Accel < 0 And pad(PadI).yVel < 0) Then
            pad(PadI).yVel = 0
            
        End If
    
    End If
     
End Sub

Private Function fnPadHitWall(PadI As Integer) As Boolean

    If (lblPaddle(PadI).Top <= 0 And pad(PadI).yVel > 0) Or (lblPaddle(PadI).Top + lblPaddle(PadI).Height > Me.ScaleHeight And pad(PadI).yVel < 0) Then
        pad(PadI).yVel = 0
        fnPadHitWall = True

    End If

End Function

Private Sub subTrails(obj As Control, UseAll As Boolean)
    
    If UseAll Then
        'NEWS DIRECTIONS
        Me.PSet (obj.Left + 10, obj.Top + (obj.Height / 2))
        Me.PSet (obj.Left + obj.Width - 10, obj.Top + (obj.Height / 2))
        Me.PSet (obj.Left + (obj.Width / 2), obj.Top + 10)
        Me.PSet (obj.Left + (obj.Width / 2), obj.Top + obj.Height - 10)
        
    End If
    
    'CORNERS
    
    Me.PSet (obj.Left + (obj.Width / 4), obj.Top + (obj.Height / 4))
    Me.PSet (obj.Left + (obj.Width * 3 / 4), obj.Top + (obj.Height / 4))
    Me.PSet (obj.Left + (obj.Width / 4), obj.Top + (obj.Height * 3 / 4))
    Me.PSet (obj.Left + (obj.Width * 3 / 4), obj.Top + (obj.Height * 3 / 4))
    
End Sub

Private Sub subResetBall()

    Randomize
    
    ball.totalVel = INIT_BALLV '150
    
    ball.xVel = Int(Rnd * (ball.totalVel - (ball.totalVel / 2)) + (ball.totalVel / 2))
    ball.yVel = fnUpdtLeg(ball.xVel, ball.totalVel)
    
    If Int(Rnd * 2) = 1 Then
        ball.yVel = -ball.yVel
    
    End If
    
    Call subLimAngle

End Sub

Private Sub subWin()
    
    Call subStop
    
    lblPrompt.Height = lblPrompt.Height * 2
    
    If game = COOP Then
        lblPrompt.Caption = "You've won!" & vbNewLine & "Please enter your duo's name."
    
    Else
        lblPrompt.Caption = "Player " & intWinner + 1 & " won!" & vbNewLine & "Please enter the winner's name."
        
    End If
    
    lblPrompt.Visible = True
    
    If pad(intWinner).AIControl = True Then
        Dim names(4) As String
        
        names(0) = "[BOT]Jerry"
        names(1) = "[BOT]Simon"
        names(2) = "[BOT]Alfred"
        names(3) = "[BOT]Jenny"
        names(4) = "[BOT]Sylvia"
        
        Randomize
        
        txtWinName.Text = names(Int(Rnd * 5))
        txtWinName.Enabled = False
    
    Else
        txtWinName.Text = "Player " & intWinner + 1 & " name"
        
    End If
            
    txtWinName.Visible = True
    
    lblEnter.Visible = True

End Sub

Private Sub subStop()

    If boolSwitch = False Then
        tmrMain.Enabled = False
        tmrRefresh.Enabled = False
        tmrPause.Enabled = False
        
        boolSwitch = True
        
    Else
        tmrMain.Enabled = True
        tmrRefresh.Enabled = True
        tmrPause.Enabled = True
        
        boolSwitch = False
        
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
     
        Case vbKeyEscape
            End
            
        Case vbKeyM
            formMenu.Show
            Unload Me
        
    End Select

End Sub

Private Sub Form_Load()
           
    'SET UP THE STAGE
    lblPaddle(1).Left = Me.ScaleWidth - (lblPaddle(0).Left + lblPaddle(0).Width)
    lblScore(1).Left = Me.ScaleWidth - (lblScore(0).Left + lblScore(0).Width)
    
    Call subCenterPos(lblEnter, Me)
    Call subCenterPos(lblPrompt, Me)
    
    Me.BackColor = vbBlack
    lblData.ForeColor = vbGreen
    Me.ForeColor = RGB(150, 255, 0)

    
    circBall.BorderColor = vbGreen
    circBall.FillColor = vbGreen
    circBall.Left = lblPaddle(0).Left + lblPaddle(0).Width + 1
    circBall.Top = Me.ScaleHeight / 2 - (circBall.Height / 2)
    
    lineMid.BorderColor = vbGreen
    lineMid.BorderStyle = 2
    lineMid.X1 = Me.ScaleWidth / 2
    lineMid.X2 = lineMid.X1
    lineMid.Y1 = 0
    lineMid.Y2 = Me.ScaleHeight
    
    lblPrompt.Left = Me.ScaleWidth / 2 - (lblPrompt.Width / 2)
    lblPrompt.ForeColor = vbGreen
    
    For i = 0 To 1
        lblPaddle(i).BackColor = vbGreen
        lblScore(i).ForeColor = vbGreen
    
    Next i
    
    Select Case game
        Case PvAI
            pad(1).AIControl = True
            
        Case AIvAI
            pad(0).AIControl = True
            pad(1).AIControl = True
            
        Case COOP
            lblScore(1).Visible = False
            
            Call subCenterPos(lblScore(0), Me)
            
            lineMid.Y1 = lblScore(0).Top + lblScore(0).Height + 100
            
    End Select
    
End Sub

Private Sub lblPause_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblPause.ForeColor = vbGreen
    lblPause.BackColor = vbBlack

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lblEnter.Visible Then
        lblEnter.ForeColor = vbBlack
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set formGame = Nothing
    
End Sub

Private Sub subUpdtHistory()

    Dim strMode As String
    Dim strContent(9) As String
    Dim i As Integer
    
    ff = FreeFile
    
    Select Case game
        Case PvP
            strMode = "PvP"
        
        Case PvAI
            strMode = "PvAI"
            
        Case AIvAI
            strMode = "DEMO"
            
        Case COOP
            strMode = "COOP"
            
    End Select
    
    Open App.Path & "\matchHistory.txt" For Input As #ff 'READ TO CLEAR OLDEST LINE

    Do
        Line Input #ff, strContent(i)
        i = i + 1
    
    Loop While Not EOF(ff) Or i = UBound(strContent)
        
    Close #ff
    
    Open App.Path & "\matchHistory.txt" For Output As #ff
    
    If game = COOP Then
        Dim WonLostGoal As String
        If lblScore(0).Caption = scoreWin Then
            WonLostGoal = "WON | " & lblScore(0).Caption & "/" & scoreWin
            
        Else
            WonLostGoal = "LOST | " & lblScore(0).Caption & "/" & scoreWin
            
        End If
        
        Print #ff, strMode & "\" & WonLostGoal & "\" & txtWinName.Text
        
    Else
        Print #ff, strMode & "\" & intScore(0) & " - " & intScore(1) & "\" & txtWinName.Text
        
    End If
    
    For j = 0 To UBound(strContent) - 1
        Print #ff, strContent(j)
    
    Next j
    
    i = 0
    Close #ff
    
End Sub

Private Sub subUpdtRecords()

    ff = FreeFile
    
    Dim strRecords() As String
    Dim intWins() As Integer
    Dim strNames() As String
    
    Dim i As Integer
    
    Open App.Path & "\records.txt" For Input As #ff
    
    ReDim strRecords(0)
    ReDim intWins(0)
    ReDim strNames(0)
    
    Do
        ReDim Preserve strRecords(i)
        ReDim Preserve intWins(i)
        ReDim Preserve strNames(i)
        
        Line Input #ff, strRecords(i)
        
        intWins(i) = Mid(strRecords(i), InStr(2, strRecords(i), "\") + 1, Len(strRecords(i)) - InStr(2, strRecords(i), "\") + 1)
        strNames(i) = Left(strRecords(i), InStr(2, strRecords(i), "\"))
        
        i = i + 1
    
    Loop While Not EOF(ff)
        
    Close #ff
    
    Open App.Path & "\records.txt" For Output As #ff
    
    Dim found As Boolean
    
    For j = 0 To UBound(strRecords)
        If strNames(j) = "\" & txtWinName.Text & "\" Then
            found = True
            intWins(j) = intWins(j) + 1
                    
        End If
    
    Next j
    
    If found = False Then
        ReDim Preserve strRecords(UBound(strRecords) + 1)
        ReDim Preserve intWins(UBound(intWins) + 1)
        ReDim Preserve strNames(UBound(strNames) + 1)
        
        strNames(UBound(strNames)) = "\" & txtWinName.Text & "\"
        intWins(UBound(intWins)) = 1
        strRecords(UBound(strRecords)) = strNames(UBound(strNames)) & intWins(UBound(intWins))
        
    End If
    
    Call subInsertSort(intWins, True, strNames)
    
    For k = 0 To UBound(strRecords)
        Print #ff, strNames(k) & intWins(k)
    
    Next k
    
    i = 0
    Close #ff
    
End Sub

Private Sub lblEnter_Click()
    
    Call subUpdtHistory
    
    '''
    
    If Not (game = COOP And lblScore(0).Caption < scoreWin) Then
        Call subUpdtRecords
        
    End If
    
    formHistory.Show
    formHistory.lblPlay.Visible = True
    
    Unload Me
    
End Sub

Private Sub lblEnter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If lblEnter.Visible Then
        lblEnter.ForeColor = vbYellow
    
    End If

End Sub

Private Sub tmrMain_Timer()

    circBall.Left = circBall.Left + ball.xVel
    circBall.Top = circBall.Top - ball.yVel
    
    For i = 0 To 1
        lblPaddle(i).Top = lblPaddle(i).Top - pad(i).yVel
        Call fnPadHitWall(Int(i))
        If Abs(pad(i).yVel) = MAX_PADV Then
            Call subTrails(lblPaddle(i), True)
          
        End If
        
    Next i
    
    If tmrPause.Enabled = False Then
        Call subHitPad
        Call subHitWall
        
    Else
        circBall.Top = fnCenter(lblPaddle(intLastHit)) - (circBall.Height / 2)
        'MsgBox 1
        
    End If
    
    If pad(0).AIControl = False Then
        Call subKeys(vbKeyW, 0, INIT_PADV)
        Call subKeys(vbKeyS, 0, -INIT_PADV)
    
    End If
    
    If (game = PvAI Or game = AIvAI) Then
        Call subAI
        
    End If
        
    If pad(1).AIControl = False Then
        Call subKeys(vbKeyUp, 1, INIT_PADV)
        Call subKeys(vbKeyDown, 1, -INIT_PADV)
        
    End If
    
    Call subCheckBounds
     
    Call subUpdtAngle
    
    If tmrPause.Enabled = False Then
        Call subTrails(circBall, False)
        
    End If
    
    lblData.Caption = "x-velocity: " & ball.xVel & ", y-velocity: " & ball.yVel & ", total velocity: " & ball.totalVel & ", angle: " & Round(dblAngle * 180 / PI, 2) & Chr(176)
    
End Sub

Private Sub tmrPause_Timer()
    
    Call subResetBall
    
    If intLastHit = 0 Then
        ball.xVel = -ball.xVel
    
    End If
    
    
    lblPrompt.Visible = False
    tmrPause.Enabled = False
    

End Sub

Private Sub tmrRefresh_Timer()

    Me.Picture = Nothing

    For i = 0 To lblPaddle.UBound
        If fnCheckCollision(circBall, lblPaddle(i)) = False Then
            lblPaddle(i).BackColor = vbGreen
            
        End If
        
    Next i

End Sub

Private Sub txtWinName_Click()

    txtWinName.Text = ""

End Sub
