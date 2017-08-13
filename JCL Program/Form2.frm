VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCurr 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstChoices 
      Height          =   1230
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtChoices 
      Height          =   1935
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Form2.lstChoices.AddItem ("JOB")
    Form2.lstChoices.AddItem ("PROC")
    For i = 0 To Form2.lstChoices.ListCount - 1
        Form2.txtChoices.Text = Form2.txtChoices.Text & Form2.lstChoices.List(i) & "     "
    Next i
End Sub
Private Sub txtChoices_Click()
    currpos = 0
    For i = 0 To Form2.lstChoices.ListCount - 1
        If Form2.txtChoices.SelStart >= currpos Then
            Form2.txtCurr.Text = Form2.lstChoices.List(i)
            currpos = currpos + Len(Form2.txtCurr.Text) + 1
            If Form2.txtChoices.SelStart < currpos Then
                Form1.Text1.SelText = Form2.txtCurr.Text
                Form1.Text1.SelStart = Form1.Text1.SelStart - Len(Form2.txtCurr.Text)
                Form1.Text1.SelLength = Len(Form2.txtCurr.Text)
                Form1.Text2.Text = Len(Form2.txtCurr.Text)
                Form2.Visible = False
                Exit For
            Else
            currpos = currpos + 4
            End If
        End If
    Next i
End Sub
