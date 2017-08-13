VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   5070
   ClientTop       =   3300
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select some text"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear selected text"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select &all text"
      Height          =   375
      Index           =   1
      Left            =   2480
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select &text from cursor"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   233
      Locked          =   -1  'True
      MousePointer    =   2  'Cross
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtSelLen As Integer
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Text1.SetFocus
    Select Case Index
        Case 0
            'Select text from cursor
            Text1.SelLength = Len(Text1.Text)
        Case 1
            'Start at beginning of text.
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
        Case 2
            'Clear selected text
            Text1.SelStart = 0
        Case 3
            'Select text immediately before
            'and immediately after cursor
            Text1.SelStart = Text1.SelStart - 3
            Text1.SelLength = 6
    End Select
End Sub

Private Sub Form_Load()
    Text2.Text = 8
End Sub

Private Sub Text1_Click()
    'Pop up a 2nd form to edit the selected text
    If Text1.SelStart > 15 And Text1.SelStart < 15 + CInt(Text2.Text) Then
        Text1.SelStart = 15
        Text1.SelLength = CInt(Text2.Text)
        Form2.Show
    End If
End Sub
