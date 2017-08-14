VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7716
   ClientLeft      =   2352
   ClientTop       =   600
   ClientWidth     =   7896
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7716
   ScaleWidth      =   7896
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   360
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      Height          =   7500
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   5000
      Begin VB.Shape block 
         Height          =   250
         Index           =   8
         Left            =   250
         Top             =   500
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   7
         Left            =   250
         Top             =   750
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   6
         Left            =   500
         Top             =   750
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   5
         Left            =   500
         Top             =   500
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   4
         Left            =   500
         Top             =   250
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   3
         Left            =   250
         Top             =   250
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   2
         Left            =   0
         Top             =   250
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   1
         Left            =   0
         Top             =   500
         Width           =   250
      End
      Begin VB.Shape block 
         Height          =   250
         Index           =   0
         Left            =   0
         Top             =   750
         Width           =   250
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If tmrSpeed.Enabled = True Then
        If KeyCode = 39 Then
        ''''Move right
            Call moveRight
        ElseIf KeyCode = 37 Then
        ''''Move left
            Call moveLeft
        ElseIf KeyCode = 40 Then
        ''''Move down faster
            tmrSpeed.Interval = tmrSpeed.Interval / 2
        ElseIf KeyCode = 90 Then
        ''''Rotate counter-clockwise
            Call rotCtrClkWse
        ElseIf KeyCode = 88 Then
        ''''Rotate clockwise
            Call rotClkWse
        End If
    End If
End Sub
Private Sub tmrSpeed_Timer()
    ''''Fall one block
    For i = 0 To block.Count - 1
        block(i).Top = block(i).Top + block(i).Height
    Next i
    ''''Test to see if hit bottom
    For i = 0 To block.Count - 1
        If block(i).Top >= Frame1.Height - block(i).Height And block(i).Visible = True Then
        ''''If hit bottom, stop fall and adjust array accordingly
            tmrSpeed.Enabled = False
        End If
    Next i
    
End Sub
Private Sub moveRight()
    For i = 0 To block.Count - 1
        If block(i).Left >= Frame1.Width - block(i).Width And block(i).Visible = True Then
            Exit Sub
        End If
    Next i
    
    For i = 0 To block.Count - 1
        block(i).Left = block(i).Left + block(i).Width
    Next i
End Sub
Private Sub moveLeft()
    For i = 0 To block.Count - 1
        If block(i).Left <= 0 And block(i).Visible = True Then
            Exit Sub
        End If
    Next i
    
    For i = 0 To block.Count - 1
        block(i).Left = block(i).Left - block(i).Width
    Next i
End Sub
Private Sub rotCtrClkWse()
    If block(0).Visible = True Then
        blockZero = True
    Else
        blockZero = False
    End If
    If block(1).Visible = True Then
        blockOne = True
    Else
        blockOne = False
    End If
    For i = 2 To 7
        If block(i).Visible = True Then
            block(i - 2).Visible = True
        Else
            block(i - 2).Visible = False
        End If
    Next i
    If blockZero = True Then
        block(6).Visible = True
    Else
        block(6).Visible = False
    End If
    If blockOne = True Then
        block(7).Visible = True
    Else
        block(7).Visible = False
    End If
End Sub
Private Sub rotClkWse()
    If block(6).Visible = True Then
        blockZero = True
    Else
        blockZero = False
    End If
    If block(7).Visible = True Then
        blockOne = True
    Else
        blockOne = False
    End If
    For i = 5 To 0 Step -1
        If block(i).Visible = True Then
            block(i + 2).Visible = True
        Else
            block(i + 2).Visible = False
        End If
    Next i
    If blockZero = True Then
        block(0).Visible = True
    Else
        block(0).Visible = False
    End If
    If blockOne = True Then
        block(1).Visible = True
    Else
        block(1).Visible = False
    End If
End Sub
