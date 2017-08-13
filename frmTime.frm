VERSION 5.00
Begin VB.Form frmTime 
   BackColor       =   &H00FFFF00&
   Caption         =   "The Time Program"
   ClientHeight    =   1695
   ClientLeft      =   1215
   ClientTop       =   3360
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   7095
   Begin VB.Timer tmrColonFlash 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2760
      Top             =   240
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   1560
      Top             =   240
   End
   Begin VB.Shape shpColon 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   135
      Index           =   2
      Left            =   4680
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape shpColon 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   135
      Index           =   3
      Left            =   4680
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   6120
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   6720
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   6120
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   6720
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   5040
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   5640
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   5040
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   5640
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   6120
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   6120
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   6120
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   5040
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   5040
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   5040
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpColon 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   135
      Index           =   1
      Left            =   2280
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape shpColon 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   135
      Index           =   0
      Left            =   2280
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   4320
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   3720
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   4320
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   3720
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   3720
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   3720
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   2640
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   2640
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   2640
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   1320
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   1320
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   1320
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpTop 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape shpMid 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape shpBot 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   3720
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   3240
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   3240
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   2640
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   2640
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   1920
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   1320
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   1920
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   1320
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpLR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   840
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpLL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape shpUR 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   840
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpUL 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gSeconds1 As Integer
Dim gSeconds10 As Integer
Dim gMinutes1 As Integer
Dim gMinutes10 As Integer
Dim gHours1 As Integer
Dim gHours10 As Integer
Private Sub Form_Load()
    gSeconds1 = frmSetTimer.getSeconds1
    gSeconds10 = frmSetTimer.getSeconds10
    gMinutes1 = frmSetTimer.getMinutes1
    gMinutes10 = frmSetTimer.getMinutes10
    gHours1 = frmSetTimer.getHours1
    gHours10 = frmSetTimer.getHours10
    Display "gSeconds1", gSeconds1
    Display "gSeconds10", gSeconds10
    Display "gMinutes1", gMinutes1
    Display "gMinutes10", gMinutes10
    Display "gHours1", gHours1
    Display "gHours10", gHours10
End Sub
Public Sub Display(X As String, Y As Integer)
    Dim indexNum As Integer

    Select Case X
        Case "gSeconds1"
            indexNum = 5
        Case "gSeconds10"
            indexNum = 4
        Case "gMinutes1"
            indexNum = 3
        Case "gMinutes10"
            indexNum = 2
        Case "gHours1"
            indexNum = 1
        Case "gHours10"
            indexNum = 0
    End Select
    
    Select Case Y
        Case 0
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = False
            shpLL(indexNum).Visible = True
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
        Case 1
            shpTop(indexNum).Visible = False
            shpUL(indexNum).Visible = False
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = False
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = False
        Case 2
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = False
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = True
            shpLR(indexNum).Visible = False
            shpBot(indexNum).Visible = True
        Case 3
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = False
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
        Case 4
            shpTop(indexNum).Visible = False
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = False
        Case 5
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = False
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
        Case 6
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = False
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = True
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
        Case 7
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = False
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = False
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = False
        Case 8
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = True
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
        Case 9
            shpTop(indexNum).Visible = True
            shpUL(indexNum).Visible = True
            shpUR(indexNum).Visible = True
            shpMid(indexNum).Visible = True
            shpLL(indexNum).Visible = False
            shpLR(indexNum).Visible = True
            shpBot(indexNum).Visible = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmSetTimer.Visible = False Then
        End
    End If
End Sub

Private Sub tmrTime_Timer()
    Display "gSeconds1", gSeconds1
    Display "gSeconds10", gSeconds10
    Display "gMinutes1", gMinutes1
    Display "gMinutes10", gMinutes10
    Display "gHours1", gHours1
    Display "gHours10", gHours10
    gSeconds1 = gSeconds1 - 1
    
    If gSeconds1 = -1 Then
        gSeconds10 = gSeconds10 - 1
        gSeconds1 = 9
        If gSeconds10 = -1 Then
            gSeconds10 = 5
            gMinutes1 = gMinutes1 - 1
            If gMinutes1 = -1 Then
                gMinutes1 = 9
                gMinutes10 = gMinutes10 - 1
                If gMinutes10 = -1 Then
                    gMinutes10 = 5
                    gHours1 = gHours1 - 1
                    If gHours1 = -1 Then
                        gHours1 = 9
                        gHours10 = gHours10 - 1
                        If gHours10 = -1 Then
                            tmrTime.Enabled = False
                        End If
                    End If
                End If
            End If
        End If
    End If
    
                        
            
    If tmrColonFlash.Enabled = False Then
        tmrColonFlash.Enabled = True
    End If
End Sub

Private Sub tmrColonFlash_Timer()
    Dim I As Integer
    For I = 0 To 3
        If shpColon(I).Visible = True Then
            shpColon(I).Visible = False
        ElseIf shpColon(I).Visible = False Then
            shpColon(I).Visible = True
        End If
    Next
    If tmrTime.Enabled = False Then
        frmTime.BackColor = &HFF00FF
        tmrColonFlash.Enabled = False
        Dim J As Integer
        For J = 0 To 3
            frmTime.shpColon(J).Visible = True
        Next
    End If
End Sub
