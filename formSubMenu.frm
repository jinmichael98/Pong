VERSION 5.00
Begin VB.Form formSubMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".pong"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10935
   Icon            =   "formSubMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox checkRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Use finer and longer trail effects"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CheckBox checkShowData 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Show ball data (velocity and angle)"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   600
      Left            =   7920
      TabIndex        =   2
      Text            =   "20"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   2280
      TabIndex        =   1
      Text            =   "3"
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "max 99"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "max 5"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Settings"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   10935
   End
   Begin VB.Label lblBack 
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
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "start game"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   4080
      TabIndex        =   8
      Top             =   5040
      Width           =   3975
   End
   Begin VB.Label lblAdjustFrict 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblAdjustFrict 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblAdjustWin 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   1
      Left            =   8640
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblAdjustWin 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   0
      Left            =   7440
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score to Win:"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image lblsDowns 
      Height          =   555
      Left            =   240
      Picture         =   "formSubMenu.frx":58142
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image lblsUps 
      Height          =   555
      Left            =   240
      Picture         =   "formSubMenu.frx":59075
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Friction"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image lblFDowns 
      Height          =   555
      Left            =   240
      Picture         =   "formSubMenu.frx":59F41
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image lblFUps 
      Height          =   555
      Left            =   240
      Picture         =   "formSubMenu.frx":5AE74
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "formSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frictFactor As Integer

Private Sub Form_Load()

    frictFactor = txtF.Text
    
    'SET STAGE
    
    Call subCenterPos(lblPlay, Me)
    Call subCenterPos(checkShowData, Me)
    Call subCenterPos(checkRef, Me)

    Label3.Left = Me.ScaleWidth / 4 - (Label3.Width / 2)
    Label2.Left = Me.ScaleWidth * 3 / 4 - (Label2.Width / 2)
    
    txtF.Left = Me.ScaleWidth / 4 - (txtF.Width / 2)
    lblAdjustFrict(0).Left = txtF.Left - lblAdjustFrict(0).Width - 200
    lblAdjustFrict(1).Left = txtF.Left + lblAdjustFrict(0).Width + 300
    
    txtScore.Left = Me.ScaleWidth * 3 / 4 - (txtScore.Width / 2)
    lblAdjustWin(0).Left = txtScore.Left - lblAdjustWin(0).Width - 200
    lblAdjustWin(1).Left = txtScore.Left + lblAdjustWin(1).Width + 300

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    For i = 0 To 1
        lblAdjustFrict(i).ForeColor = vbGreen
        lblAdjustWin(i).ForeColor = vbGreen
    
    Next i
    
    lblPlay.ForeColor = vbBlack
    
    lblBack.ForeColor = vbBlack
    
End Sub

Private Sub lblAdjustFrict_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If frictFactor > 1 Then
                txtF.Text = Int(txtF.Text) - 1
        
            End If
        
        Case 1
            If frictFactor <= 5 Then
                txtF.Text = Int(txtF.Text) + 1
    
            End If
    
    End Select
    
End Sub

Private Sub lblAdjustFrict_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblAdjustFrict(Index).ForeColor = vbYellow

End Sub

Private Sub lblAdjustWin_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If txtScore.Text > 1 Then
                txtScore.Text = Int(txtScore.Text) - 1
        
            End If
        
        Case 1
            If txtScore.Text < 99 Then
                txtScore.Text = Int(txtScore.Text) + 1
    
            End If
    
    End Select
    
    
End Sub

Private Sub lblAdjustWin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblAdjustWin(Index).ForeColor = vbYellow

End Sub

Private Sub lblBack_Click()
    
    formMenu.Show
    Unload Me

End Sub

Private Sub lblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblBack.ForeColor = vbYellow

End Sub

Private Sub lblPlay_Click()
    
    friction = frictFactor * 0.15
    
    scoreWin = Int(txtScore.Text)
        
    formGame.Show
        
    If checkShowData.Value = 0 Then
        formGame.lblData.Visible = False
            
    Else
        formGame.lblData.Visible = True
            
    End If
    
    If checkRef.Value = 1 Then
        formGame.tmrRefresh.Interval = 200
        formGame.DrawWidth = 1
        
    Else
        formGame.DrawWidth = 2
    
    End If

    Unload Me
        
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblPlay.ForeColor = vbYellow

End Sub

Private Sub txtF_Change()
    
    If IsNumeric(txtF.Text) Then
        If txtF.Text > 5 Then
            txtF.Text = "5"
            
        ElseIf txtF.Text < 1 Then
            txtF.Text = "1"
        
        End If
        
        frictFactor = txtF.Text
        
        If IsNumeric(txtScore.Text) Then
            lblPlay.Enabled = True
        
        End If
        
    Else: lblPlay.Enabled = False
    
    End If
    
End Sub

Private Sub txtScore_Change()

    If IsNumeric(txtScore.Text) Then
        If txtScore.Text > 99 Then
        
            txtScore.Text = "99"
            
        ElseIf txtScore.Text < 1 Then
        
            txtScore.Text = "1"
        
        End If
        
        If IsNumeric(txtF.Text) Then
            lblPlay.Enabled = True
        
        End If
        
    Else
        lblPlay.Enabled = False
    End If

End Sub
