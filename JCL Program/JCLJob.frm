VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Statement"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "PROC"
      Top             =   660
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "JCL"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label a 
      Caption         =   "Choose the statement type:"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Form1.Text3.Text = Text4.Text
    Form1.Text4.Text = Text5.Text
    Form2.Visible = False
End Sub

Private Sub Text2_Click()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
    Text5.SelText = Text2.Text
End Sub
Private Sub Text3_Click()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
    Text5.SelText = Text3.Text
End Sub
