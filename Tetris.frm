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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   990
      Top             =   1110
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   990
      Top             =   855
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   990
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   735
      Top             =   1110
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   735
      Top             =   855
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   735
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   480
      Top             =   1110
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   480
      Top             =   855
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gShapeNum

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If gShapeNum < 3 Then
                gShapeNum = gShapeNum + 1
            Else
                gShapeNum = 0
            End If
            If Timer1.Enabled = True Then
                ChangeShape
            End If
        Case vbKeyLeft
            If Timer1.Enabled = True And Shape1(0).Left >= Shape1(0).Width Then
                Dim I As Integer
                For I = 0 To 8
                    Shape1(I).Left = Shape1(I).Left - Shape1(I).Width
                Next
            End If
        Case vbKeyRight
            If Timer1.Enabled = True And Shape1(8).Left <= Form1.Width - Shape1(8).Width Then
                Dim A As Integer
                For A = 0 To 8
                    Shape1(A).Left = Shape1(A).Left + Shape1(A).Width
                Next
            End If
        End Select
End Sub

Private Sub Form_Load()
    gShapeNum = 0
    ChangeShape
End Sub

Public Sub ChangeShape()
    Select Case gShapeNum
        Case 0
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            Shape1(2).Visible = True
            Shape1(3).Visible = False
            Shape1(4).Visible = False
            Shape1(5).Visible = True
            Shape1(6).Visible = False
            Shape1(7).Visible = False
            Shape1(8).Visible = False
        Case 1
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            Shape1(2).Visible = False
            Shape1(3).Visible = True
            Shape1(4).Visible = False
            Shape1(5).Visible = False
            Shape1(6).Visible = True
            Shape1(7).Visible = False
            Shape1(8).Visible = False
        Case 2
            Shape1(0).Visible = False
            Shape1(1).Visible = False
            Shape1(2).Visible = False
            Shape1(3).Visible = True
            Shape1(4).Visible = False
            Shape1(5).Visible = False
            Shape1(6).Visible = True
            Shape1(7).Visible = True
            Shape1(8).Visible = True
        Case 3
            Shape1(0).Visible = False
            Shape1(1).Visible = True
            Shape1(2).Visible = True
            Shape1(3).Visible = False
            Shape1(4).Visible = False
            Shape1(5).Visible = True
            Shape1(6).Visible = False
            Shape1(7).Visible = False
            Shape1(8).Visible = True
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim I As Integer
    For I = 0 To 8
        Shape1(I).Top = Shape1(I).Top + Shape1(I).Height
    Next
    If Shape1(8).Top >= Form1.Height - 3 * Shape1(8).Height Then
        Timer1.Enabled = False
    End If
End Sub
