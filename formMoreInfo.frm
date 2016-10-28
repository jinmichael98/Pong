VERSION 5.00
Begin VB.Form formMoreInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".pong"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   FillColor       =   &H0000FF00&
   Icon            =   "formMoreInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10695
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
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "formMoreInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    lblInfo.Caption = _
    ":: Pong 2.0 ::" & vbNewLine _
     & vbNewLine _
     & "-A player's OBJECTIVE is to gain enough POINTS to win" & vbNewLine & vbNewLine _
     & "-You can manipulate the ball's movement by allowing it to bounce off your paddle" & vbNewLine & vbNewLine _
     & "-In non-coop game modes, POINTS are earned by sending the ball past the enemy paddle(s)" & vbNewLine & vbNewLine _
     & "-In the coop game mode, POINTS are earned for every ROTATION of the ball between the paddles" & vbNewLine & vbNewLine _
     & "-Paddles have a MAXIMUM speed and they give off a TRAIL when they are moving at this speed" & vbNewLine & vbNewLine _
     & "-Unlike in classic Pong, the ball's movement is affected by your paddle's SPEED ON IMPACT, not by location" & vbNewLine & vbNewLine _
     & "-The ball will BOUNCE off the far edges of a paddle, so try to have it hit the MIDDLE of your paddle" & vbNewLine & vbNewLine _
     & "-SERVES send the ball to a RANDOM direction unlike normal bounces you choose where to serve by moving" & vbNewLine & vbNewLine _
     & "-FRICTION affects how much the ball's movement channges with your paddle's speed" & vbNewLine & vbNewLine _
     & "-Computers play better in low-friction environments, you can use this to change game DIFFICULTY" & vbNewLine & vbNewLine _
     & "-To become the Pong Champion, you must accumulate the MOST WINS recorded" & vbNewLine _

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbBlack

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set formMoreInfo = Nothing

End Sub

Private Sub lblMenu_Click()

    formMenu.Show
    Unload Me

End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbYellow

End Sub
