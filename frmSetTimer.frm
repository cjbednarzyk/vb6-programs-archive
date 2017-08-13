VERSION 5.00
Begin VB.Form frmSetTimer 
   Caption         =   "Timer Initialization"
   ClientHeight    =   2640
   ClientLeft      =   2745
   ClientTop       =   630
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   2940
   Begin VB.Timer tmrDelayMenu 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   1440
   End
   Begin VB.TextBox txtSeeMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdBeginCountDown 
      Caption         =   "&Begin Countdown!!!"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ComboBox cboSeconds 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cboMinutes 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cboHours 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds:"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes:"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours:"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Timer"
      Visible         =   0   'False
      Begin VB.Menu mnuBegCount 
         Caption         =   "&Begin Countdown"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFeatures 
      Caption         =   "&Features"
      Visible         =   0   'False
      Begin VB.Menu mnuCommandButtonsEnabled 
         Caption         =   "&Command Buttons Enabled"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHideSetTimer 
         Caption         =   "&Hide This Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSetTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gHours10 As Integer
Dim gHours1 As Integer
Dim gMinutes10 As Integer
Dim gMinutes1 As Integer
Dim gSeconds10 As Integer
Dim gSeconds1 As Integer

Private Sub cboHours_Change()
    Dim I As Integer
    Dim J As Boolean
    
    For I = 0 To 24
        If cboHours.Text = I Then
            J = True
        End If
    Next
    
    If J = False Then
        cboHours.Text = 0
        cboHours.SetFocus
    End If
End Sub

Private Sub cboMinutes_Change()
    Dim I As Integer
    Dim J As Boolean
    
    For I = 0 To 59
        If cboMinutes.Text = I Then
            J = True
        End If
    Next
    
    If J = False Then
        cboMinutes.Text = 0
        cboMinutes.SetFocus
    End If
End Sub

Private Sub cboSeconds_Change()
    Dim I As Integer
    Dim J As Boolean
    
    For I = 0 To 59
        If cboSeconds.Text = I Then
            J = True
        End If
    Next
    
    If J = False Then
        cboSeconds.Text = 0
        cboSeconds.SetFocus
    End If
End Sub

Private Sub cmdBeginCountDown_Click()
    Unload frmTime
    If cmdPause.Caption = "&Continue" Then
        cmdPause.Caption = "&Pause"
        mnuPause.Caption = "&Pause"
    End If
    
    If cboHours.Text >= 20 Then
        gHours10 = 2
        gHours1 = cboHours.Text - 20
    ElseIf cboHours.Text >= 10 Then
        gHours10 = 1
        gHours1 = cboHours.Text - 10
    Else
        gHours10 = 0
        gHours1 = cboHours.Text
    End If
    
    If cboMinutes.Text >= 50 Then
        gMinutes10 = 5
        gMinutes1 = cboMinutes.Text - 50
    ElseIf cboMinutes.Text >= 40 Then
        gMinutes10 = 4
        gMinutes1 = cboMinutes.Text - 40
    ElseIf cboMinutes.Text >= 30 Then
        gMinutes10 = 3
        gMinutes1 = cboMinutes.Text - 30
    ElseIf cboMinutes.Text >= 20 Then
        gMinutes10 = 2
        gMinutes1 = cboMinutes.Text - 20
    ElseIf cboMinutes.Text >= 10 Then
        gMinutes10 = 1
        gMinutes1 = cboMinutes.Text - 10
    Else
        gMinutes10 = 0
        gMinutes1 = cboMinutes.Text
    End If
     
    If cboSeconds.Text >= 50 Then
        gSeconds10 = 5
        gSeconds1 = cboSeconds.Text - 50
    ElseIf cboSeconds.Text >= 40 Then
        gSeconds10 = 4
        gSeconds1 = cboSeconds.Text - 40
    ElseIf cboSeconds.Text >= 30 Then
        gSeconds10 = 3
        gSeconds1 = cboSeconds.Text - 30
    ElseIf cboSeconds.Text >= 20 Then
        gSeconds10 = 2
        gSeconds1 = cboSeconds.Text - 20
    ElseIf cboSeconds.Text >= 10 Then
        gSeconds10 = 1
        gSeconds1 = cboSeconds.Text - 10
    ElseIf cboSeconds.Text >= 1 Then
        gSeconds10 = 0
        gSeconds1 = cboSeconds.Text
    End If
    
    If cboSeconds.Text = 0 Then
        If cboMinutes.Text = 0 Then
            If cboHours.Text = 0 Then
            Else
                frmTime.Show
            End If
        Else
                frmTime.Show
        End If
    Else
            frmTime.Show
    End If
    'MsgBox Str(gSeconds1) + " " + Str(gSeconds10) + " " + Str(gMinutes1) + " " + Str(gMinutes10) + " " + Str(gHours1) + " " + Str(gHours10)
    
    
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdPause_Click()
    If cmdPause.Caption = "&Pause" Then
        cmdPause.Caption = "&Continue"
        mnuPause.Caption = "&Continue"
        frmTime.tmrTime.Enabled = False
        frmTime.tmrColonFlash.Enabled = False
        Dim I As Integer
        For I = 0 To 3
            frmTime.shpColon(I).Visible = True
        Next
            
    ElseIf cmdPause.Caption = "&Continue" Then
        cmdPause.Caption = "&Pause"
        mnuPause.Caption = "&Pause"
        frmTime.tmrTime.Enabled = True
        frmTime.tmrColonFlash.Enabled = True
    End If
End Sub
Private Sub Form_Load()
    frmSplash.Visible = True
    frmSetTimer.Visible = False
    Dim I As Integer
    For I = 0 To 24
        cboHours.AddItem I
    Next
    Dim J As Integer
    For J = 0 To 59
        cboMinutes.AddItem J
        cboSeconds.AddItem J
    Next
End Sub
Public Function getSeconds1()
    getSeconds1 = gSeconds1
End Function
Public Function getSeconds10()
    getSeconds10 = gSeconds10
End Function
Public Function getMinutes1()
    getMinutes1 = gMinutes1
End Function
Public Function getMinutes10()
    getMinutes10 = gMinutes10
End Function
Public Function getHours1()
    getHours1 = gHours1
End Function
Public Function getHours10()
    getHours10 = gHours10
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTime
End Sub
Private Sub mnuBegCount_Click()
    cmdBeginCountDown_Click
    MenuView
    End Sub

Private Sub mnuCommandButtonsEnabled_Click()
    If mnuCommandButtonsEnabled.Checked = True Then
        cmdExit.Visible = False
        cmdBeginCountDown.Visible = False
        cmdPause.Visible = False
        mnuCommandButtonsEnabled.Checked = False
        frmSetTimer.Height = 1720
    Else
        cmdExit.Visible = True
        cmdBeginCountDown.Visible = True
        cmdPause.Visible = True
        mnuCommandButtonsEnabled.Checked = True
        frmSetTimer.Height = 3045
    End If
    MenuView
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHideSetTimer_Click()
    If frmTime.Visible = True Then
        MenuView
        frmSetTimer.Visible = False
    Else
        Dim cool
        cool = MsgBox("You must exit if you wish to hide this window.  Is this okay?", vbOKCancel, "Exit?")
        If cool = vbOK Then
            End
        End If
    End If
End Sub

Private Sub mnuPause_Click()
    cmdPause_Click
    MenuView
End Sub


Private Sub tmrDelayMenu_Timer()
    tmrDelayMenu.Enabled = False
    If mnuProgram.Visible = False Then
        mnuProgram.Visible = True
        mnuHelp.Visible = True
        mnuFeatures.Visible = True
        cmdExit.Top = 1680
        cmdPause.Top = 1680
        cmdBeginCountDown.Top = 1200
        cboHours.Top = 720
        cboMinutes.Top = 720
        cboSeconds.Top = 720
        lblHours.Top = 240
        lblMinutes.Top = 240
        lblSeconds.Top = 240
        txtSeeMenu.Top = -400
    Else
        mnuProgram.Visible = False
        mnuHelp.Visible = False
        mnuFeatures.Visible = False
        cmdExit.Top = 1920
        cmdPause.Top = 1920
        cmdBeginCountDown.Top = 1440
        cboHours.Top = 960
        cboMinutes.Top = 960
        cboSeconds.Top = 960
        lblHours.Top = 480
        lblMinutes.Top = 480
        lblSeconds.Top = 480
        txtSeeMenu.Top = 0
    End If
End Sub

Private Sub txtSeeMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuView
End Sub

Public Sub MenuView()
    tmrDelayMenu.Enabled = True
End Sub
