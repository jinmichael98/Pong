VERSION 5.00
Begin VB.Form formHistory 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ".pong"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10965
   Icon            =   "formHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2430
      Index           =   2
      Left            =   6480
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox lstHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2430
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox lstHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2430
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "play again"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblBest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Most Wins:"
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
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   10965
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
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Winner"
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
      Index           =   2
      Left            =   6480
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
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
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "formHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TextLine
    strLine As String
    mode As String
    result As String
    winner As String
    position As Integer

End Type

Dim i As Integer
Dim lines(9) As TextLine

Private Sub subEnterData()

    ff = FreeFile
    
    Open App.Path & "\matchHistory.txt" For Input As #ff
    
    Do
        Input #ff, lines(i).strLine
        
        For j = 1 To Len(lines(i).strLine)
            If Mid(lines(i).strLine, j, 1) = "\" Then
                lines(i).position = lines(i).position + 1
                
            Else
                Select Case lines(i).position
                    Case 0 'mode
                        lines(i).mode = lines(i).mode & Mid(lines(i).strLine, j, 1)
                        
                    Case 1 'res
                        lines(i).result = lines(i).result & Mid(lines(i).strLine, j, 1)
                    
                    Case 2 'win
                        lines(i).winner = lines(i).winner & Mid(lines(i).strLine, j, 1)
                        
                End Select
            
            End If

        Next j
                
        lines(i).position = 0
                    
        lstHistory(0).AddItem lines(i).mode
        lstHistory(1).AddItem lines(i).result
        lstHistory(2).AddItem lines(i).winner
        
        i = i + 1
    
    Loop Until EOF(ff) Or i = UBound(lines) + 1
    
    Close #ff

End Sub

Private Sub subMostWin()

    ff = FreeFile
    
    Dim strBest As String
    Dim strBestScore As String
    Dim strBestName As String
    Dim place As Integer
    
    Open App.Path & "\records.txt" For Input As #ff
    Line Input #ff, strBest
    
    For j = 2 To Len(strBest)
        If Mid(strBest, j, 1) = "\" Then
            place = place + 1
        
        Else
            Select Case place
                Case 0
                    strBestName = strBestName + Mid(strBest, j, 1)
                    
                Case 1
                    strBestScore = strBestScore + Mid(strBest, j, 1)
            
            End Select
            
        End If
        
    Next j
    
    Close #ff
    
    lblBest.Caption = "Pong Champion: " & strBestName & " @ " & strBestScore & " wins"

End Sub

Private Sub Form_Load()

   Call subEnterData
   
   Call subMostWin
   
   Call subCenterPos(lblPlay, Me)
   

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbBlack
    lblPlay.ForeColor = vbBlack

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set formHistory = Nothing

End Sub

Private Sub lblMenu_Click()

    formMenu.Show
    
    Unload Me

End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblMenu.ForeColor = vbYellow

End Sub

Private Sub lblPlay_Click()

    formGame.Show
    Unload Me

End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblPlay.ForeColor = vbYellow

End Sub
