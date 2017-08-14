VERSION 5.00
Begin VB.Form frmTicTacToeStart 
   Caption         =   "Tic Tac Toe Startup"
   ClientHeight    =   3552
   ClientLeft      =   5508
   ClientTop       =   2952
   ClientWidth     =   3132
   LinkTopic       =   "Form1"
   ScaleHeight     =   3552
   ScaleWidth      =   3132
   Begin VB.Frame frameNumPlayers 
      Height          =   1092
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2892
      Begin VB.OptionButton optTwoPlayers 
         Caption         =   "Two Players"
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1812
      End
      Begin VB.OptionButton optOnePlayer 
         Caption         =   "Vs. Computer"
         Height          =   372
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1692
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Please select    X or O to start    "
      Height          =   732
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   1212
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   696
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   696
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   696
   End
End
Attribute VB_Name = "frmTicTacToeStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblBox_Click(Index As Integer)
    If Index = 0 Then
        frmTicTacToe.turn = "X"
        frmTicTacToe.lblPlayer1.ForeColor = vbRed
        frmTicTacToe.lblPlayer1.Caption = "X"
        frmTicTacToe.lblPlayer2.ForeColor = vbBlue
        frmTicTacToe.lblPlayer2.Caption = "O"
    Else
        frmTicTacToe.turn = "O"
        frmTicTacToe.lblPlayer1.ForeColor = vbBlue
        frmTicTacToe.lblPlayer1.Caption = "O"
        frmTicTacToe.lblPlayer2.ForeColor = vbRed
        frmTicTacToe.lblPlayer2.Caption = "X"
    End If
    frmTicTacToe.lblInfo.Caption = "It is now " + frmTicTacToe.turn + "'s turn!"
    Unload Me
End Sub

Private Sub optOnePlayer_Click()
    frmTicTacToe.onePlayer = True
End Sub

Private Sub optTwoPlayers_Click()
    frmTicTacToe.onePlayer = False
End Sub
