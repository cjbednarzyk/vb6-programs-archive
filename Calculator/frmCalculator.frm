VERSION 5.00
Begin VB.Form frmCalculator 
   Caption         =   "Christopher's Calculator Program"
   ClientHeight    =   5640
   ClientLeft      =   4140
   ClientTop       =   1725
   ClientWidth     =   6180
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleMode       =   0  'User
   ScaleWidth      =   6180
   Begin VB.CommandButton cmdList 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3315
      TabIndex        =   21
      ToolTipText     =   "List all previous calculations"
      Top             =   1200
      Width           =   750
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   1320
      TabIndex        =   18
      ToolTipText     =   "Delete the last number in the present entry"
      Top             =   1200
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   4320
      TabIndex        =   20
      ToolTipText     =   "Cancel all previous calculations (Reset Calculator)"
      Top             =   1200
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Index           =   1
      Left            =   5325
      TabIndex        =   19
      ToolTipText     =   "Cancel the present entry"
      Top             =   2100
      Width           =   750
   End
   Begin VB.TextBox txtMem 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   26
      ToolTipText     =   "Tells if there's a number in memory"
      Top             =   1200
      Width           =   750
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   200
      TabIndex        =   23
      ToolTipText     =   "Store Into Memory"
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   200
      TabIndex        =   25
      ToolTipText     =   "Clear Memory"
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   3
      Left            =   200
      TabIndex        =   22
      ToolTipText     =   "Add into memory"
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdMem 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   200
      TabIndex        =   24
      ToolTipText     =   "Recall Memory"
      Top             =   3000
      Width           =   750
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5325
      TabIndex        =   17
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   4
      Left            =   5325
      TabIndex        =   16
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   4320
      TabIndex        =   15
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   4320
      TabIndex        =   14
      Top             =   3000
      Width           =   750
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   4320
      TabIndex        =   13
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3315
      TabIndex        =   11
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1320
      TabIndex        =   10
      Top             =   4800
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   3
      Left            =   3315
      TabIndex        =   3
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   2
      Left            =   2325
      TabIndex        =   2
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   3900
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   6
      Left            =   3315
      TabIndex        =   6
      Top             =   3000
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   5
      Left            =   2325
      TabIndex        =   5
      Top             =   3000
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   9
      Left            =   3315
      TabIndex        =   9
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   8
      Left            =   2325
      TabIndex        =   8
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   2100
      Width           =   750
   End
   Begin VB.CommandButton cmdNum 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   2325
      TabIndex        =   0
      Top             =   4800
      Width           =   750
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   200
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   200
      Width           =   5895
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
   Begin VB.Menu mnuSplash 
      Caption         =   "&View Splash Screen"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op1, Op2 As Double 'Previously input operand
Dim Memory As Double 'Number in Memory
Dim InMemory As Boolean 'Is there a number in memory?
Dim TempReadout  'Temporary stored value
Dim DecimalFlag As Integer 'Is the decimal point present yet?
Dim NumOps As Integer 'How many operands?
Dim LastInput As String 'What was the last type of event?
Dim OpFlag As String 'What type of operation is pending?

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim I As Integer
    For I = 0 To 9
        If Chr(KeyAscii) = CStr(I) Then
            cmdNum(I).Value = True
            GoTo LastLine
        End If
    Next
    If Chr(KeyAscii) = "%" Then
        cmdPercent.Value = True
    ElseIf Chr(KeyAscii) = "*" Then
        cmdOp(0).Value = True
    ElseIf Chr(KeyAscii) = "/" Then
        cmdOp(1).Value = True
    ElseIf Chr(KeyAscii) = "-" Then
        cmdOp(2).Value = True
    ElseIf Chr(KeyAscii) = "+" Then
        cmdOp(3).Value = True
    ElseIf Chr(KeyAscii) = "=" Then
        cmdOp(4).Value = True
    ElseIf Chr(KeyAscii) = "C" Then
        cmdCancel(0).Value = True
    ElseIf Chr(KeyAscii) = "E" Then
        cmdCancel(1).Value = True
    ElseIf Chr(KeyAscii) = "B" Then
        cmdCancel(2).Value = True
    ElseIf Chr(KeyAscii) = "L" Then
        cmdList.Value = True
    ElseIf Chr(KeyAscii) = "." Then
        cmdDecimal.Value = True
    ElseIf Chr(KeyAscii) = "N" Then
        cmdPlusMinus.Value = True
    ElseIf Chr(KeyAscii) = "M" Then
        cmdMem(0).Value = True
    ElseIf Chr(KeyAscii) = "R" Then
        cmdMem(1).Value = True
    ElseIf Chr(KeyAscii) = "S" Then
        cmdMem(2).Value = True
    ElseIf Chr(KeyAscii) = "P" Then
        cmdMem(3).Value = True
    End If
LastLine:
End Sub
Private Sub Form_Load()
    Memory = 0
    InMemory = False
    DecimalFlag = False
    NumOps = 0
    LastInput = ""
    OpFlag = ""
    txtDisplay = ""
    txtMem = ""
End Sub
Private Sub cmdCancel_Click(Index As Integer)
    Select Case cmdCancel(Index).Caption
        Case "C"
            Op1 = Op2 = 0
            Dim B As Boolean
            If frmListOps.Visible = True Then
                B = True
            Else
                B = False
            End If
            Form_Load
            frmListOps.lstOps.Clear
            If B = False Then
                frmListOps.Visible = False
            End If
        Case "CE"
            txtDisplay = frmListOps.lstOps.List(frmListOps.lstOps.ListCount - 1)
            DecimalFlag = False
            LastInput = "CE"
            NumOps = 1
        Case "Backspace"
            If txtDisplay <> "" Then
                txtDisplay = Left(txtDisplay, Len(txtDisplay) - 1)
            End If
    End Select
End Sub
Private Sub cmdPlusMinus_Click()
    If LastInput <> "OPS" Then
        If txtDisplay = "-" Then
            txtDisplay = ""
        ElseIf Left(txtDisplay, 1) = "-" Then
            txtDisplay = Right(txtDisplay, Len(txtDisplay) - 1)
        ElseIf txtDisplay = "" Then
            txtDisplay = "-"
        Else
            txtDisplay = "-" + txtDisplay
        End If
    Else
        txtDisplay = "-"
        DecimalFlag = False
    End If
    LastInput = "NUMS"
End Sub
Private Sub cmdPercent_Click()
    txtDisplay = txtDisplay / 100
    LastInput = "OPS"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub
Private Sub cmdDecimal_Click()
        If LastInput = "NEG" Then
            txtDisplay = "-0."
        ElseIf LastInput <> "NUMS" Then
            txtDisplay = "0."
        ElseIf DecimalFlag = False Then
            txtDisplay = txtDisplay + "."
        End If
        DecimalFlag = True
        LastInput = "NUMS"
End Sub
Private Sub cmdNum_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        txtDisplay = ""
        DecimalFlag = False
    End If
    If LastInput = "NEG" Then
        txtDisplay = "-" & txtDisplay
    End If
    txtDisplay = txtDisplay + CStr(Index)
    LastInput = "NUMS"
End Sub
Private Sub cmdOp_Click(Index As Integer)
    TempReadout = txtDisplay
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
            If cmdOp(Index).Caption = "-" And LastInput <> "NEG" Then
                txtDisplay = "-" & txtDisplay
                LastInput = "NEG"
            End If
        Case 1
            Op1 = txtDisplay
            If cmdOp(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
                txtDisplay = "-"
                LastInput = "NEG"
            End If
            If frmListOps.lstOps.List(frmListOps.lstOps.ListCount - 1) <> Op1 Then
                frmListOps.lstOps.AddItem (Op1)
            End If
        Case 2
            Op2 = TempReadout
            Select Case OpFlag
                Case "+"
                    Op1 = CDbl(Op1) + CDbl(Op2)
                Case "-"
                    Op1 = CDbl(Op1) - CDbl(Op2)
                Case "*"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                Case "/"
                    If Op2 = 0 Then
                        MsgBox "Can't Divide By Zero!", 48, "Calculator"
                    Else
                        Op1 = CDbl(Op1) / CDbl(Op2)
                    End If
                Case "="
                    Op1 = CDbl(Op2)
                Case "%"
                    Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
            txtDisplay = Op1
            NumOps = 1
            Select Case OpFlag
                Case "="
                    If frmListOps.lstOps.List(frmListOps.lstOps.ListCount - 1) <> Op1 Then
                        frmListOps.lstOps.AddItem ("=====")
                        frmListOps.lstOps.AddItem (Op1)
                    End If
                Case Else
                    frmListOps.lstOps.AddItem (OpFlag)
                    frmListOps.lstOps.AddItem (Op2)
                    frmListOps.lstOps.AddItem ("----------")
                    frmListOps.lstOps.AddItem (Op1)
            End Select
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = cmdOp(Index).Caption
    End If
End Sub
Private Sub cmdMem_Click(Index As Integer)
    Select Case cmdMem(Index).Caption
        Case "MS"
            If txtDisplay <> "" Then
                Memory = CDbl(txtDisplay)
                InMemory = True
                txtMem = "M"
            End If
        Case "M+"
            If InMemory Then
                Memory = Memory + CDbl(txtDisplay)
            ElseIf txtDisplay <> "" Then
                Memory = CDbl(txtDisplay)
                InMemory = True
                txtMem = "M"
            End If
        Case "MC"
            InMemory = False
            txtMem = ""
        Case "MR"
            If InMemory Then
                DecimalFlag = False
                txtDisplay = Memory
                LastInput = "NUMS"
            End If
    End Select
End Sub
Private Sub cmdList_Click()
    If frmListOps.Visible = False Then
        frmListOps.Visible = True
    Else
        frmListOps.Visible = False
    End If
End Sub
Private Sub mnuClose_Click()
    End
End Sub
Private Sub Form_Resize()
    'Maintain the same form proportions
    'frmCalculator.Width = frmCalculator.Height * 6270 / 6330
    'frmCalculator.Height = frmCalculator.Width * 6330 / 6270
    'Adjust the display's size/location properties
    With txtDisplay
        .Top = 200 / 6330 * frmCalculator.Height
        .Height = 750 / 6330 * frmCalculator.Height
        .Left = 200 / 6270 * frmCalculator.Width
        .Width = 5895 / 6270 * frmCalculator.Width
    End With
    'Adjust the number keys' size/location properties
    Dim I As Integer
    For I = 0 To 9
        cmdNum(I).Height = 750 / 6330 * frmCalculator.Height
        cmdNum(I).Width = 750 / 6270 * frmCalculator.Width
    Next
    For I = 1 To 7 Step 3
        cmdNum(I).Left = 1320 / 6270 * frmCalculator.Width
    Next
    cmdNum(0).Left = 2325 / 6270 * frmCalculator.Width
    For I = 2 To 8 Step 3
        cmdNum(I).Left = 2325 / 6270 * frmCalculator.Width
    Next
    For I = 3 To 9 Step 3
        cmdNum(I).Left = 3315 / 6270 * frmCalculator.Width
    Next
    cmdNum(0).Top = 4800 / 6330 * frmCalculator.Height
    For I = 1 To 3
        cmdNum(I).Top = 3900 / 6330 * frmCalculator.Height
    Next
    For I = 4 To 6
        cmdNum(I).Top = 3000 / 6330 * frmCalculator.Height
    Next
    For I = 7 To 9
        cmdNum(I).Top = 2100 / 6330 * frmCalculator.Height
    Next
    'Adjust the memory keys' size/location properties
    For I = 0 To 3
        With cmdMem(I)
            .Top = ((2100 / 6330) + (900 * I) / 6330) * frmCalculator.Height
            .Left = 200 / 6270 * frmCalculator.Width
            .Width = cmdNum(I).Width 'Being lazy :)
            .Height = cmdNum(I).Height 'Again!
        End With
    Next
    With txtMem
        .Width = cmdMem(0).Width
        .Height = cmdMem(0).Height
        .Top = cmdMem(0).Top - 900 / 6330 * frmCalculator.Height
        .Left = cmdMem(0).Left
    End With
    'Take care of the operator keys
    For I = 0 To 3
        With cmdOp(I)
            .Top = cmdMem(I).Top 'Being lazy
            .Left = 4320 / 6270 * frmCalculator.Width
            .Width = cmdMem(I).Width
            .Height = cmdMem(I).Height
        End With
    Next
    With cmdOp(4)
        .Top = cmdOp(3).Top
        .Height = cmdOp(3).Height
        .Width = cmdOp(3).Width
        .Left = 5325 / 6270 * frmCalculator.Width
    End With
    'Percent Key
    With cmdPercent
        .Height = cmdOp(4).Height
        .Width = cmdOp(4).Width
        .Left = cmdOp(4).Left
        .Top = 3900 / 6330 * frmCalculator.Height
    End With
    'CE Key
    With cmdCancel(1)
        .Width = cmdPercent.Width
        .Left = cmdPercent.Left
        .Top = cmdOp(0).Top
        .Height = (1650 / 6330) * frmCalculator.Height
    End With
    'C, Backspace Keys
    For I = 0 To 2 Step 2
        With cmdCancel(I)
            .Height = cmdNum(I).Height
            .Top = (1200 / 6330) * frmCalculator.Height
            .Width = (1755 / 6270) * frmCalculator.Width
            .Left = ((4320 / 6270) - (1500 * I) / 6270) * frmCalculator.Width
        End With
    Next
    'List Key
    With cmdList
        .Height = txtMem.Height
        .Width = txtMem.Width
        .Top = txtMem.Top
        .Left = cmdNum(9).Left
    End With
    'Plus/minus key
    With cmdPlusMinus
        .Height = cmdNum(0).Height
        .Width = cmdNum(0).Width
        .Top = cmdNum(0).Top
        .Left = (1320 / 6270) * frmCalculator.Width
    End With
    'Decimal Key
    With cmdDecimal
        .Height = cmdPlusMinus.Height
        .Width = cmdPlusMinus.Width
        .Top = cmdPlusMinus.Top
        .Left = (3315 / 6270) * frmCalculator.Width
    End With
End Sub

Private Sub mnuSplash_Click()
    frmSplash.Visible = True
End Sub
