VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Job/Proc Statement"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   350
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Show
End Sub
Private Sub Text2_Click()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text2_DblClick()
    Form2.Show
End Sub
Private Sub Text3_Click()
    If Text3.Text = "" Then
        Text3.SelText = "<none>"
    End If
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text3_DblClick()
    Form2.Show
End Sub
Private Sub Text4_Click()
    If Text4.Text = "" Then
        Text4.SelText = "<none>"
    End If
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text4_DblClick()
    Form2.Show
End Sub
