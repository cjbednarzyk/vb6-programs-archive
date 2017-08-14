VERSION 5.00
Begin VB.Form frmTicTacToe 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4680
   ClientLeft      =   3504
   ClientTop       =   2256
   ClientWidth     =   7224
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7224
   Begin VB.Frame framePlayer2 
      Caption         =   "Player 2"
      Height          =   1092
      Left            =   5280
      TabIndex        =   18
      Top             =   2280
      Width           =   1692
      Begin VB.Label lblPlayer2 
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   696
      End
      Begin VB.Label lblScore2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   28.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "&Undo Move"
      Enabled         =   0   'False
      Height          =   492
      Left            =   2880
      TabIndex        =   11
      Top             =   3840
      Width           =   1092
   End
   Begin VB.CommandButton cmdPlyAgn 
      Caption         =   "&Play Again"
      Height          =   492
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Height          =   852
      Left            =   1440
      TabIndex        =   12
      Top             =   3600
      Width           =   2652
   End
   Begin VB.Frame framePlayer1 
      Caption         =   "Player 1"
      Height          =   1092
      Left            =   5280
      TabIndex        =   13
      Top             =   1080
      Width           =   1692
      Begin VB.Label lblScore1 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   28.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   612
      End
      Begin VB.Label lblPlayer1 
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   696
      End
   End
   Begin VB.Frame frameWins 
      Caption         =   "Total Wins"
      Height          =   2652
      Left            =   5160
      TabIndex        =   15
      Top             =   840
      Width           =   1932
   End
   Begin VB.Label Label1 
      Height          =   732
      Left            =   6240
      TabIndex        =   17
      Top             =   1800
      Width           =   612
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   6
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   7
      Left            =   2400
      TabIndex        =   9
      Top             =   2640
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   8
      Left            =   3120
      TabIndex        =   10
      Top             =   2640
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   1920
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   4
      Left            =   2400
      TabIndex        =   6
      Top             =   1920
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   696
   End
   Begin VB.Label lblBox 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   696
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   696
   End
   Begin VB.Label lblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2532
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public turn As String 'X or O?
Public onePlayer As Boolean 'If true, one player.  If false, two players.
Private i As Integer  'generic counter - reused throughout
Private j As Integer 'same as i
Private win As Boolean 'Did the last move result in a win?
Private loaded As Boolean 'Is the subform loaded?  If so, don't load it.
Private movenum As Variant 'Which move is this?
Private history(9, 2) As Variant 'There are up to 9 moves in tic-tac-toe.  Each move consists of a location and a value (X or O).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private currView(9) As String 'The current view of the tic-tac-toe board.
Private currMax As Integer 'The current maximum evaluation function value.
Private currCount As Integer 'The current evaluation function value.
Private currMove As Integer 'The current best move for the computer to make.
Private yourTurn As String
Private yourPotentialWins As Integer
Private myPotentialWins As Integer
Private myWin As Boolean
Private yourWin As Boolean

Private Sub cmdPlyAgn_Click()
    For i = 0 To 8
        lblBox(i).Caption = ""
    Next i
    movenum = 1
    win = False
    cmdUndo = False
    lblInfo.Caption = "It is now " + turn + "'s turn!"
End Sub
Private Sub cmdUndo_Click()
    movenum = movenum - 1
    lblBox(history(movenum, 2)) = ""
    turn = history(movenum, 1)
    lblInfo.Caption = "It is now " + turn + "'s turn!"
    If movenum <= 1 Then
        cmdUndo.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    If loaded = False Then
        frmTicTacToeStart.Show
        frmTicTacToeStart.SetFocus
        loaded = True
    End If
End Sub

Private Sub Form_Load()
    Randomize
    loaded = False
    movenum = 1
End Sub

Private Sub lblBox_Click(Index As Integer)
    If lblBox(Index).Caption = "" Then
        cmdUndo.Enabled = True
        history(movenum, 2) = Index
        If turn = "X" Then
            lblBox(Index).ForeColor = vbRed
            lblBox(Index).Caption = "X"
            history(movenum, 1) = "X"
            turn = "O"
        Else
            lblBox(Index).ForeColor = vbBlue
            lblBox(Index).Caption = "O"
            history(movenum, 1) = "O"
            turn = "X"
        End If
        movenum = movenum + 1
        lblInfo.Caption = "It is now " + turn + "'s turn!"
        Call winOccurred(Index)
        If win = True Then
            cmdUndo.Enabled = False
            If lblBox(Index).Caption = lblPlayer1.Caption Then
                lblScore1.Caption = CStr(CInt(lblScore1.Caption) + 1)
            Else
                lblScore2.Caption = CStr(CInt(lblScore2.Caption) + 1)
            End If
            lblInfo.Caption = lblBox(Index).Caption + " Wins!"
            Exit Sub
        End If
        If movenum = 10 Then
            cmdUndo.Enabled = False
            lblInfo.Caption = "Nobody Won!"
        End If
        If onePlayer = True And turn <> history(1, 1) Then
            Call makeMove
        End If
    Else
        lblInfo.Caption = "It is now " + turn + "'s turn!"
        lblInfo.Caption = lblInfo.Caption + "Try again!"
    End If
End Sub
Private Sub winOccurred(Index As Integer)
    'Test the row for a win
    win = True
    For i = (Index \ 3) * 3 To ((Index \ 3) * 3 + 2)
        If lblBox(i).Caption <> lblBox(Index).Caption Then win = False
    Next i
    If win = True Then Exit Sub
    'Test the column for a win
    win = True
    For i = (Index Mod 3) To 8 Step 3
        If lblBox(i).Caption <> lblBox(Index).Caption Then win = False
    Next i
    If win = True Then Exit Sub
    'Test the upper-left to bottom-right diagonal for a win
    win = True
    For i = 0 To 8 Step 4
        If lblBox(i).Caption <> lblBox(Index).Caption Then win = False
    Next i
    If win = True Then Exit Sub
    'Test the bottom-left to upper-right diagonal for a win
    win = True
    For i = 2 To 6 Step 2
        If lblBox(i).Caption <> lblBox(Index).Caption Then win = False
    Next i
End Sub
Private Sub makeMove()
    currMove = Int(Rnd * (9 - movenum))
    For i = 8 To 0
        If lblBox(i).Caption = "" Then
            currMove = currMove - 1
            If currMove <= 0 Then
                currMove = i
                Exit For
            End If
        End If
    Next i
    Call lblBox_Click(currMove)
End Sub
