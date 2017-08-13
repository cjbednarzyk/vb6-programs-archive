VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   1080
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Timer tmrInformUser 
         Interval        =   300
         Left            =   2760
         Top             =   2640
      End
      Begin VB.Label lblInformUser 
         Caption         =   "Initializing Timer Window..."
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label lblYear 
         Caption         =   "February 4, 1999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5010
         TabIndex        =   5
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCompany 
         Caption         =   "Christopher J. Bednarzyk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5010
         TabIndex        =   1
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "For Windows 95"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5010
         TabIndex        =   2
         Top             =   2700
         Width           =   1845
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4995
         TabIndex        =   3
         Top             =   2340
         Width           =   1860
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "The TIMER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2355
         TabIndex        =   4
         Top             =   705
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gCount As Integer

Private Sub Form_Load()
    gCount = 0
End Sub
Private Sub tmrInformUser_Timer()
    Select Case gCount
        Case 0
            lblInformUser.Caption = "Setting up menu..."
        Case 1
            lblInformUser.Caption = "Initializing Combo Box Setup..."
        Case 2
            lblInformUser.Caption = "Initializing Timer..."
        Case 3
            lblInformUser.Caption = "Program Starting..."
        Case 6
            frmSplash.Visible = False
        Case 8
            tmrInformUser.Enabled = False
            Unload frmSplash
            frmSetTimer.Visible = True
    End Select
    gCount = gCount + 1
End Sub
