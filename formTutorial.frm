VERSION 5.00
Begin VB.Form formTutorial 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".pong"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formTutorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMain 
      Interval        =   10
      Left            =   480
      Top             =   3360
   End
   Begin VB.Timer tmrTutorial 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblKeyInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Go to Menu"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   15
      Top             =   4995
      Width           =   1815
   End
   Begin VB.Label lblM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9720
      TabIndex        =   14
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblKeyInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End program"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   4995
      Width           =   1815
   End
   Begin VB.Label lblEsc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Esc"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player Controls"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1845
      TabIndex        =   11
      Top             =   5760
      Width           =   7335
   End
   Begin VB.Label lblMore 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "More Game Info"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9360
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblKeyInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "move down"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblKeyInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "move up"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPaddle 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7080
      TabIndex        =   6
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      X1              =   7560
      X2              =   7560
      Y1              =   3840
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      X1              =   3240
      X2              =   3240
      Y1              =   3840
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      X1              =   3240
      X2              =   7560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   3240
      X2              =   7560
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Image imgDown 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   6480
      Picture         =   "formTutorial.frx":58142
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Image imgUp 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   6480
      Picture         =   "formTutorial.frx":58413
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label lblW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   2  'Dash
      X1              =   5520
      X2              =   5520
      Y1              =   360
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblPaddle 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "formTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intVel(1) As Integer

Private Sub subKeys(CheckKey As Long, PadI As Integer, Accel As Integer) 'KEY FOR VBKEY VALUE AND PADI FOR INDEX, ACCEL FOR VALUE OF ACCELERATION
    
    If GetAsyncKeyState(CheckKey) And ((Accel > 0 And lblPaddle(PadI).Top > 0) Or (Accel < 0 And lblPaddle(PadI).Top + lblPaddle(PadI).Height < Me.ScaleHeight)) And _
    ((lblPaddle(PadI).Top > Label4.Top And Accel > 0) Or (lblPaddle(PadI).Top + lblPaddle(PadI).Height < Label4.Top + Label4.Height And Accel < 0)) Then   'LIMIT MAXIMUM VELOCITY
        If intVel(PadI) < 200 Then
            intVel(PadI) = intVel(PadI) + Accel
        
        End If
        
    Else
        If (Accel > 0 And intVel(PadI) > 0) Or (Accel < 0 And intVel(PadI) < 0) Then
            intVel(PadI) = 0
            
        End If
            
    End If
    
End Sub

Private Sub Form_Load()

    lblM.Left = Me.ScaleWidth - (lblEsc.Left + lblEsc.Width)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbBlack
    lblMore.ForeColor = vbBlack

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set formTutorial = Nothing

End Sub

Private Sub lblMenu_Click()

    formMenu.Show
    Unload Me

End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbYellow

End Sub

Private Sub lblMore_Click()

    formMoreInfo.Show
    Unload Me

End Sub

Private Sub lblMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMore.ForeColor = vbYellow
    
End Sub

Private Sub tmrMain_Timer()

    Call subKeys(vbKeyW, 0, 20)
    Call subKeys(vbKeyS, 0, -20)
    
    Call subKeys(vbKeyUp, 1, 20)
    Call subKeys(vbKeyDown, 1, -20)
    
    For i = 0 To 1
        lblPaddle(i).Top = lblPaddle(i).Top - intVel(i)
        
        If lblPaddle(i).Top < Label4.Top Then
            lblPaddle(i).Top = Label4.Top
        
        ElseIf lblPaddle(i).Top + lblPaddle(i).Height > Label4.Top + Label4.Height Then
            lblPaddle(i).Top = Label4.Top + Label4.Height - lblPaddle(i).Height
            
        End If
        
    Next i

End Sub
